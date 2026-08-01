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
    private static readonly StrategyBasedComWrappers Wrappers = new();
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
    /// Takes ownership of a raw automation pointer. The pointer must already be AddRef'd for us,
    /// which is the case for anything COM hands to a callback or returns from a call.
    /// </summary>
    public static DispatchObject? Attach(nint pointer)
    {
        if (pointer == 0)
        {
            return null;
        }

        var managed = Wrappers.GetOrCreateObjectForComInstance(pointer, CreateObjectFlags.UniqueInstance);
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
            VarEnum.VT_EMPTY or VarEnum.VT_NULL => 0,
            _ => Convert.ToInt32(value.As<object>(), System.Globalization.CultureInfo.InvariantCulture),
        };
    }

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
                0,
                0);

            if (hr < 0)
            {
                result.Dispose();
                Marshal.ThrowExceptionForHR(hr);
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
