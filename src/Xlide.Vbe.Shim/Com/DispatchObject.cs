using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// Owns one automation object obtained from the editor and calls it by member name.
///
/// Ownership is single and explicit: an instance releases exactly the reference it was created
/// with, exactly once, and objects it hands back are owned by their new holder. Disposing a parent
/// does not dispose a child that was already returned to a caller.
///
/// Calls go through dispatch rather than a compiled vtable. For control-plane work, which happens
/// once per user action or per editor event, the cost is irrelevant and it removes any dependency
/// on the member ordering of a particular typelib build. Hot paths that run per keystroke use
/// early-bound interfaces instead.
/// </summary>
internal sealed unsafe class DispatchObject : IDisposable
{
    private static readonly Guid NullGuid = Guid.Empty;

    private nint _pointer;
    private IDispatch? _dispatch;

    private DispatchObject(nint pointer, IDispatch dispatch)
    {
        _pointer = pointer;
        _dispatch = dispatch;
    }

    /// <summary>True when this wrapper still holds a live reference.</summary>
    public bool IsAlive => _pointer != 0;

    /// <summary>
    /// The underlying automation pointer, borrowed. Callers pass it into another automation call
    /// and never release it; this wrapper remains the owner.
    /// </summary>
    public nint Pointer => _pointer;

    /// <summary>
    /// Takes ownership of a raw automation pointer. The pointer must already be AddRef'd for us,
    /// which is the case for anything COM hands to a callback or returns from a call.
    /// </summary>
    public static DispatchObject? Attach(nint pointer)
    {
        if (pointer == 0)
        {
            return null;
        }

        var managed = ComRuntime.Wrappers.GetOrCreateObjectForComInstance(pointer, CreateObjectFlags.UniqueInstance);
        if (managed is not IDispatch dispatch)
        {
            Marshal.Release(pointer);
            return null;
        }

        return new DispatchObject(pointer, dispatch);
    }

    /// <summary>
    /// Adds a reference to a borrowed pointer and wraps it. Use when the caller keeps its own
    /// reference, such as a pointer received as an inbound call argument.
    /// </summary>
    public static DispatchObject? AttachBorrowed(nint pointer)
    {
        if (pointer == 0)
        {
            return null;
        }

        Marshal.AddRef(pointer);
        return Attach(pointer);
    }

    /// <summary>Resolves a member name to its dispatch identifier.</summary>
    public int GetDispId(string name)
    {
        ObjectDisposedException.ThrowIf(_dispatch is null, this);

        fixed (char* namePointer = name)
        {
            var namePointers = stackalloc nint[1];
            namePointers[0] = (nint)namePointer;

            var dispIds = stackalloc int[1];
            dispIds[0] = DispId.Unknown;

            var hr = _dispatch.GetIDsOfNames(NullGuid, (nint)namePointers, 1, 0, (nint)dispIds);
            return hr < 0 ? DispId.Unknown : dispIds[0];
        }
    }

    /// <summary>Reads a property and returns its raw variant. The caller disposes the result.</summary>
    public ComVariant GetProperty(string name)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        return InvokeCore(dispId, InvokeKind.PropertyGet, []);
    }

    /// <summary>Reads a property as text, or null when it is absent or empty.</summary>
    public string? GetString(string name)
    {
        using var value = GetProperty(name);
        return VariantToString(value);
    }

    /// <summary>Reads a property as a 32-bit integer.</summary>
    public int GetInt32(string name)
    {
        using var value = GetProperty(name);
        return value.VarType switch
        {
            VarEnum.VT_I4 => value.As<int>(),
            VarEnum.VT_I2 => value.As<short>(),
            VarEnum.VT_INT => value.As<int>(),
            VarEnum.VT_BOOL => value.As<bool>() ? 1 : 0,
            VarEnum.VT_EMPTY or VarEnum.VT_NULL => 0,

            // Named rather than coerced. Reading a variant as object is not supported and throws
            // "UnsupportedType", which says nothing about which member was read or what it held:
            // a boolean property read through here failed with exactly that and looked like a
            // marshalling fault rather than the wrong reader for the type.
            _ => throw new InvalidOperationException(
                $"'{name}' is a {value.VarType} and cannot be read as a number."),
        };
    }

    /// <summary>
    /// Reads a property as a boolean.
    ///
    /// Automation booleans are not the same width as a language boolean and true is not 1, so this
    /// goes through the variant's own conversion rather than comparing a number.
    /// </summary>
    public bool GetBool(string name)
    {
        using var value = GetProperty(name);
        return value.VarType switch
        {
            VarEnum.VT_BOOL => value.As<bool>(),
            VarEnum.VT_I4 => value.As<int>() != 0,
            VarEnum.VT_I2 => value.As<short>() != 0,
            VarEnum.VT_EMPTY or VarEnum.VT_NULL => false,
            _ => throw new InvalidOperationException(
                $"'{name}' is a {value.VarType} and cannot be read as a boolean."),
        };
    }

    /// <summary>
    /// Reads a property that takes two numbers and returns text. This is how a code module hands
    /// over a range of lines, which is the only way to read a module's source in one call rather
    /// than one call per line.
    /// </summary>
    public string? GetStringIndexed(string name, int first, int second)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        using var a = ComVariant.Create(first);
        using var b = ComVariant.Create(second);
        using var value = InvokeCore(dispId, InvokeKind.PropertyGet | InvokeKind.Method, [a, b]);

        return VariantToString(value);
    }

    /// <summary>Writes a numeric property.</summary>
    public void SetInt32(string name, int value) => SetProperty(name, ComVariant.Create(value));

    /// <summary>Writes a floating-point property.</summary>
    public void SetDouble(string name, double value) => SetProperty(name, ComVariant.Create(value));

    /// <summary>The variant type a property currently holds, which is how its kind is learned.</summary>
    public VarEnum GetVarType(string name)
    {
        using var value = GetProperty(name);
        return value.VarType;
    }

    /// <summary>
    /// Reads a property once and reports both the type it held and how it prints.
    ///
    /// One read, deliberately. A property is read by running its getter, a getter is code that can
    /// do anything, and some do a great deal: reading a workbook's mail session starts the mail
    /// system. Callers that need both answers must not pay that price twice.
    /// </summary>
    public (VarEnum Kind, string Display) ReadProperty(string name)
    {
        using var value = GetProperty(name);
        return (value.VarType, Display(value));
    }

    /// <summary>Writes a text property.</summary>
    public void SetString(string name, string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        SetProperty(name, ComVariant.Create(value));
    }

    /// <summary>
    /// Calls a method that takes another automation object, which is how a collection is asked to
    /// remove one of its members.
    /// </summary>
    public void InvokeWithObject(string name, DispatchObject argument)
    {
        ArgumentNullException.ThrowIfNull(argument);
        ObjectDisposedException.ThrowIf(!argument.IsAlive, argument);

        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        // The variant releases what it holds when it is cleared, so it is given a reference of its
        // own rather than the caller's. Without this the argument is released twice.
        Marshal.AddRef(argument.Pointer);

        using var value = ComVariant.CreateRaw(VarEnum.VT_DISPATCH, argument.Pointer);
        using var result = InvokeCore(dispId, InvokeKind.Method, [value]);
    }

    /// <summary>Calls a method that takes one number and returns another automation object.</summary>
    public DispatchObject? CallObject(string name, int argument)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        using var value = ComVariant.Create(argument);
        using var result = InvokeCore(dispId, InvokeKind.Method | InvokeKind.PropertyGet, [value]);
        return FromVariant(result);
    }

    /// <summary>
    /// Calls a method with one string argument and renders whatever it returns as text.
    ///
    /// The result is shown to a developer, so it is rendered the way the language spells things
    /// rather than the way this runtime does: True rather than true, Empty and Null by name.
    /// </summary>
    public string CallToString(string name, string argument)
    {
        ArgumentNullException.ThrowIfNull(argument);

        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        using var value = ComVariant.Create(argument);
        using var result = InvokeCore(dispId, InvokeKind.Method | InvokeKind.PropertyGet, [value]);
        return Display(result);
    }

    /// <summary>Renders a variant the way VBA would print it.</summary>
    private static string Display(in ComVariant value) => value.VarType switch
    {
        VarEnum.VT_EMPTY => "Empty",
        VarEnum.VT_NULL => "Null",
        VarEnum.VT_BOOL => value.As<bool>() ? "True" : "False",
        VarEnum.VT_I2 => value.As<short>().ToString(CultureInfo.InvariantCulture),
        VarEnum.VT_I4 or VarEnum.VT_INT => value.As<int>().ToString(CultureInfo.InvariantCulture),
        VarEnum.VT_R4 => value.As<float>().ToString(CultureInfo.InvariantCulture),
        VarEnum.VT_R8 => value.As<double>().ToString(CultureInfo.InvariantCulture),
        VarEnum.VT_DATE => value.As<DateTime>().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        VarEnum.VT_BSTR => VariantToString(value) ?? string.Empty,
        VarEnum.VT_DISPATCH or VarEnum.VT_UNKNOWN => "[object]",
        _ => value.VarType.ToString(),
    };

    /// <summary>Reads a property that returns another automation object.</summary>
    public DispatchObject? GetObject(string name)
    {
        using var value = GetProperty(name);
        return FromVariant(value);
    }

    /// <summary>Calls a method that returns an object, with no arguments.</summary>
    public DispatchObject? CallObject(string name)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        using var value = InvokeCore(dispId, InvokeKind.Method | InvokeKind.PropertyGet, []);
        return FromVariant(value);
    }

    /// <summary>Calls a method that takes numbers and returns another automation object.</summary>
    public DispatchObject? CallObject(string name, params ReadOnlySpan<int> arguments)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        // Numbers own nothing, so there is nothing to release per argument.
        var variants = new ComVariant[arguments.Length];
        for (var i = 0; i < arguments.Length; i++)
        {
            variants[i] = ComVariant.Create(arguments[i]);
        }

        using var value = InvokeCore(dispId, InvokeKind.Method | InvokeKind.PropertyGet, variants);
        return FromVariant(value);
    }

    /// <summary>Calls a method that returns nothing and takes no arguments.</summary>
    public void Invoke(string name)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        using var result = InvokeCore(dispId, InvokeKind.Method, []);
    }

    /// <summary>
    /// Calls a method that returns nothing and takes numbers, which is every positioning call on
    /// the editor object model.
    /// </summary>
    public void Invoke(string name, params ReadOnlySpan<int> arguments)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        // Numbers own nothing, so there is no per-argument cleanup here. Anything that allocates
        // (a string, an object) must be released through the array element rather than a copy of
        // it, which is why this overload is deliberately limited to numbers.
        var variants = new ComVariant[arguments.Length];
        for (var i = 0; i < arguments.Length; i++)
        {
            variants[i] = ComVariant.Create(arguments[i]);
        }

        using var result = InvokeCore(dispId, InvokeKind.Method, variants);
    }

    /// <summary>
    /// Calls a method whose arguments are all numbers the callee writes into.
    ///
    /// This is how the editor reports where the caret is: it has no property for it, only a method
    /// with four out parameters. The storage has to be ours and has to outlive the call, and each
    /// descriptor points straight at one of the slots rather than at a second variant, because a
    /// variant of a variant would only reach the callee through the dispatch layer's coercion and
    /// whether that runs depends on how the callee implements Invoke.
    /// </summary>
    public void InvokeInt32s(string name, Span<int> results)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        fixed (int* values = results)
        {
            var variants = new ComVariant[results.Length];
            for (var i = 0; i < results.Length; i++)
            {
                variants[i] = ComVariant.CreateRaw(VarEnum.VT_BYREF | VarEnum.VT_I4, (nint)(values + i));
            }

            // Nothing is freed for a by-reference variant: the storage is the caller's, which is
            // this stack frame.
            using var result = InvokeCore(dispId, InvokeKind.Method, variants);
        }
    }

    /// <summary>
    /// Calls a method that takes one string and returns nothing, which is how a module is given
    /// its source.
    /// </summary>
    public void Invoke(string name, string argument)
    {
        ArgumentNullException.ThrowIfNull(argument);

        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        // The variant owns the string it was built from and frees it when it is cleared, so it is
        // disposed here rather than after the call returns.
        using var value = ComVariant.Create(argument);
        using var result = InvokeCore(dispId, InvokeKind.Method, [value]);
    }

    /// <summary>
    /// Reads an indexed member, which is how automation collections expose their items. Editor
    /// collections are one-based.
    /// </summary>
    public DispatchObject? GetItem(string collectionMember, int index)
    {
        using var collection = GetObject(collectionMember);
        return collection?.GetItem(index);
    }

    /// <summary>Reads item <paramref name="index"/> from a collection object.</summary>
    public DispatchObject? GetItem(int index)
    {
        var dispId = GetDispId("Item");
        if (dispId == DispId.Unknown)
        {
            dispId = DispId.Value;
        }

        using var argument = ComVariant.Create(index);
        using var value = InvokeCore(dispId, InvokeKind.Method | InvokeKind.PropertyGet, [argument]);
        return FromVariant(value);
    }

    /// <summary>Reads a named item from a collection that indexes by name as well as position.</summary>
    public DispatchObject? GetItem(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);

        var dispId = GetDispId("Item");
        if (dispId == DispId.Unknown)
        {
            dispId = DispId.Value;
        }

        // The variant owns the string it was built from and frees it when disposed.
        using var argument = ComVariant.Create(name);
        using var value = InvokeCore(dispId, InvokeKind.Method | InvokeKind.PropertyGet, [argument]);
        return FromVariant(value);
    }

    /// <summary>
    /// Calls a method whose last parameter is an in-out automation object, and reports both the
    /// method's return value and the object the callee wrote into that parameter.
    ///
    /// An out parameter cannot be expressed as a return value, so the storage has to be ours, has
    /// to outlive the call, and has to be read afterwards. The descriptor points straight at an
    /// interface slot rather than at a second variant, because that is how the parameter is
    /// declared: a byref variant would reach the callee only through the dispatch layer's type
    /// coercion, and whether that coercion runs depends on how the callee implements Invoke.
    /// </summary>
    public DispatchObject? CallWithByRefObject(
        string name,
        ReadOnlySpan<ComVariant> leadingArguments,
        out DispatchObject? byRefResult)
    {
        var dispId = GetDispId(name);
        if (dispId == DispId.Unknown)
        {
            throw new InvalidOperationException($"The object has no member named '{name}'.");
        }

        nint slot = 0;

        var arguments = new ComVariant[leadingArguments.Length + 1];
        leadingArguments.CopyTo(arguments);
        arguments[^1] = ComVariant.CreateRaw(VarEnum.VT_BYREF | VarEnum.VT_DISPATCH, (nint)(&slot));

        using var value = InvokeCore(dispId, InvokeKind.Method, arguments);

        // Whatever the callee put in the slot, it counted for us.
        byRefResult = Attach(slot);
        return FromVariant(value);
    }

    /// <summary>Writes a boolean property.</summary>
    public void SetBool(string name, bool value) => SetProperty(name, ComVariant.Create(value));

    /// <summary>
    /// Writes a property that takes another automation object, which is how the active pane is
    /// chosen. The variant is given a reference of its own, because clearing it releases one.
    ///
    /// Assigned by reference, which is its own invoke kind: an object property put that arrives
    /// as an ordinary put asks the callee to copy a value, and the editor refuses that outright.
    /// </summary>
    public void SetObject(string name, DispatchObject value)
    {
        ArgumentNullException.ThrowIfNull(value);
        ObjectDisposedException.ThrowIf(!value.IsAlive, value);

        Marshal.AddRef(value.Pointer);
        SetProperty(name, ComVariant.CreateRaw(VarEnum.VT_DISPATCH, value.Pointer), InvokeKind.PropertyPutRef);
    }

    /// <summary>
    /// Writes any property.
    ///
    /// A property assignment is the one call shape that carries a named argument: the value being
    /// assigned is identified by a reserved dispatch identifier rather than by position. Without it
    /// the call arrives with a value the callee cannot account for and is refused.
    ///
    /// This is worth stating plainly because it was got wrong once and the wrong conclusion was
    /// drawn from it. Two setters here passed the value positionally, every assignment through them
    /// failed, and the failures were read as the editor refusing to allow those properties to be
    /// set. They were refusals of a malformed call.
    /// </summary>
    private void SetProperty(string name, ComVariant value, InvokeKind kind = InvokeKind.PropertyPut)
    {
        ObjectDisposedException.ThrowIf(_dispatch is null, this);

        try
        {
            var dispId = GetDispId(name);
            if (dispId == DispId.Unknown)
            {
                throw new InvalidOperationException($"The object has no member named '{name}'.");
            }

            var namedArgument = DispId.PropertyPut;
            var parameters = default(DispatchParameters);
            parameters.Arguments = (nint)(&value);
            parameters.ArgumentCount = 1;
            parameters.NamedArguments = (nint)(&namedArgument);
            parameters.NamedArgumentCount = 1;

            var exception = default(ExcepInfo);

            var hr = _dispatch.Invoke(
                dispId,
                NullGuid,
                0,
                (ushort)kind,
                (nint)(&parameters),
                0,
                (nint)(&exception),
                0);

            if (hr < 0)
            {
                Throw(hr, ref exception);
            }
        }
        finally
        {
            value.Dispose();
        }
    }

    private static DispatchObject? FromVariant(in ComVariant value)
    {
        if (value.VarType is not (VarEnum.VT_DISPATCH or VarEnum.VT_UNKNOWN))
        {
            return null;
        }

        var pointer = value.GetRawDataRef<nint>();
        if (pointer == 0)
        {
            return null;
        }

        // The variant still owns its reference and will release on dispose, so take our own.
        return AttachBorrowed(pointer);
    }

    private static string? VariantToString(in ComVariant value) => value.VarType switch
    {
        VarEnum.VT_BSTR => value.As<string?>(),
        VarEnum.VT_EMPTY or VarEnum.VT_NULL => null,
        _ => value.As<object>()?.ToString(),
    };

    /// <summary>DISP_E_EXCEPTION: the callee raised an error and filled in the description.</summary>
    private const int DispatchException = unchecked((int)0x80020009);

    /// <summary>
    /// Raises the failure, preferring what the callee said over what the runtime would say.
    ///
    /// The strings in the block belong to the caller once it has been filled in, so they are freed
    /// here whether or not they are used.
    /// </summary>
    private static void Throw(int hr, ref ExcepInfo exception)
    {
        try
        {
            if (hr != DispatchException)
            {
                Marshal.ThrowExceptionForHR(hr);
                return;
            }

            var description = exception.Description == 0 ? null : Marshal.PtrToStringBSTR(exception.Description);
            var source = exception.Source == 0 ? null : Marshal.PtrToStringBSTR(exception.Source);

            if (string.IsNullOrWhiteSpace(description))
            {
                Marshal.ThrowExceptionForHR(exception.ErrorCode != 0 ? exception.ErrorCode : hr);
                return;
            }

            throw new InvalidOperationException(
                string.IsNullOrWhiteSpace(source) ? description : $"{description.Trim()}");
        }
        finally
        {
            FreeBstr(ref exception.Source);
            FreeBstr(ref exception.Description);
            FreeBstr(ref exception.HelpFile);
        }
    }

    private static void FreeBstr(ref nint value)
    {
        if (value != 0)
        {
            Marshal.FreeBSTR(value);
            value = 0;
        }
    }

    private ComVariant InvokeCore(int dispId, InvokeKind kind, ReadOnlySpan<ComVariant> arguments)
    {
        ObjectDisposedException.ThrowIf(_dispatch is null, this);

        var result = default(ComVariant);
        var parameters = default(DispatchParameters);

        // Automation expects arguments in reverse order.
        var reversed = arguments.Length == 0 ? [] : new ComVariant[arguments.Length];
        for (var i = 0; i < arguments.Length; i++)
        {
            reversed[i] = arguments[arguments.Length - 1 - i];
        }

        // Rich error information is asked for and used.
        //
        // When a call fails because the callee raised an error, the interesting part is the
        // description the callee wrote, not the HRESULT: "Type mismatch" against "Arg_COMException".
        // Passing nothing here throws the second one away, which for anything the developer typed
        // means the answer is discarded and a generic wrapper shown in its place.
        var exception = default(ExcepInfo);

        fixed (ComVariant* argumentBlock = reversed)
        {
            parameters.Arguments = arguments.Length == 0 ? 0 : (nint)argumentBlock;
            parameters.ArgumentCount = (uint)arguments.Length;

            var hr = _dispatch.Invoke(
                dispId,
                NullGuid,
                0,
                (ushort)kind,
                (nint)(&parameters),
                (nint)(&result),
                (nint)(&exception),
                0);

            if (hr < 0)
            {
                result.Dispose();
                Throw(hr, ref exception);
            }
        }

        return result;
    }

    public void Dispose()
    {
        var pointer = Interlocked.Exchange(ref _pointer, 0);
        _dispatch = null;

        if (pointer != 0)
        {
            Marshal.Release(pointer);
        }
    }
}
