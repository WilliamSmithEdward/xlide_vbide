#if DEBUG
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Shim.Com;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The api answered from INSIDE the host process: the object a live session registers in the
/// Running Object Table as "Xlide.Api", so code already running - a workbook's VBA, or an
/// automation client - reaches the same routes the HTTP door serves:
///
///     GetObject(, "Xlide.Api").Request("agent")
///     GetObject(, "Xlide.Api").Request("module?name=Module1")
///     GetObject(, "Xlide.Api").Request("module?name=Scratch", source)
///
/// This is the most direct path the api has: the call arrives on the host thread, the route runs
/// right there (CrossToHost recognises its own thread and goes inline), and the answer returns
/// as one BSTR - no socket, no marshaling beyond the two strings. What a caller gives up is
/// everything that needs the host thread to PUMP while waiting - the page-script routes - which
/// the policy in AgentGuide refuses fast with a pointer at the HTTP door.
///
/// THE DOOR IS MECHANISM ONLY. Route policy - what answers here, what is HTTP-only - lives with
/// the session's delegate; this class parses two strings, calls, and marshals one string back.
/// It implements IDispatch by hand the way <see cref="XlideAddIn"/> does, because VBA reaches
/// everything late-bound.
/// </summary>
[GeneratedComClass]
internal sealed partial class InsideDoor : IDispatch
{
    private const int RequestDispId = 1;
    private const int GuideDispId = 2;

    private readonly Func<string, string, string> _answer;

    /// <param name="answer">(route-with-query, body) to the reply JSON. Policy included: the
    /// session decides what an inside caller may reach, not this object.</param>
    public InsideDoor(Func<string, string, string> answer)
    {
        _answer = answer;
    }

    public int GetTypeInfoCount(out uint count)
    {
        count = 0;
        return HResult.Ok;
    }

    public int GetTypeInfo(uint typeInfoIndex, uint lcid, out nint typeInfo)
    {
        typeInfo = 0;
        return HResult.Fail;
    }

    public unsafe int GetIDsOfNames(in Guid riid, nint names, uint nameCount, uint lcid, nint dispIds)
    {
        if (names == 0 || dispIds == 0 || nameCount == 0)
        {
            return HResult.InvalidArg;
        }

        var namePointers = (char**)names;
        var results = (int*)dispIds;
        var resolvedAll = true;

        for (var i = 0u; i < nameCount; i++)
        {
            // Case-insensitive, unlike the extensibility interface above: the callers here are
            // people typing into the Immediate window, not a host reading a vtable.
            var name = Marshal.PtrToStringUni((nint)namePointers[i]);
            var dispId = name?.ToUpperInvariant() switch
            {
                "REQUEST" => RequestDispId,
                "GUIDE" => GuideDispId,
                _ => DispId.Unknown,
            };

            results[i] = dispId;
            if (dispId == DispId.Unknown)
            {
                resolvedAll = false;
            }
        }

        return resolvedAll ? HResult.Ok : HResult.DispUnknownName;
    }

    public unsafe int Invoke(
        int dispIdMember,
        in Guid riid,
        uint lcid,
        ushort flags,
        nint dispParams,
        nint result,
        nint exceptionInfo,
        nint argumentError)
    {
        if (dispParams == 0)
        {
            return HResult.InvalidArg;
        }

        // Assignments are refused, not executed. A property PUT delivers its value as the one
        // argument, in the same slot a call's argument arrives in - so without this guard,
        // `door.Request = "state"` would RUN the state route as a side effect of an assignment
        // statement, which is the least expected thing an assignment could do.
        if ((flags & (ushort)(InvokeKind.PropertyPut | InvokeKind.PropertyPutRef)) != 0)
        {
            return HResult.DispMemberNotFound;
        }

        var parameters = (DispatchParameters*)dispParams;
        var arguments = (ComVariant*)parameters->Arguments;
        var count = parameters->ArgumentCount;

        // Automation passes arguments in reverse order.
        ComVariant* Argument(uint indexFromLeft) =>
            arguments is null || indexFromLeft >= count ? null : &arguments[count - 1 - indexFromLeft];

        string answer;
        switch (dispIdMember)
        {
            case RequestDispId when count >= 1:
            {
                var target = StringOf(Argument(0));
                if (target is null)
                {
                    answer = """{"error":"Request takes a route string, e.g. Request(\"agent\")"}""";
                    break;
                }

                var body = count >= 2 ? StringOf(Argument(1)) ?? string.Empty : string.Empty;
                answer = Answer(target, body);
                break;
            }

            case GuideDispId:
                // The property spelling of "introduce yourself": the same reply `agent` serves.
                answer = Answer("agent", string.Empty);
                break;

            default:
                return HResult.DispMemberNotFound;
        }

        if (result != 0)
        {
            // The BSTR is allocated here and owned by the caller from this line on, which is the
            // automation contract for a return value.
            *(ComVariant*)result = ComVariant.Create(answer);
        }

        return HResult.Ok;
    }

    /// <summary>The delegate, behind the one promise this object makes: never throw across COM.
    /// An exception surfacing from Invoke would reach VBA as an unhelpful automation error; the
    /// api's own convention - an error field in the JSON - survives the crossing intact.</summary>
    private string Answer(string target, string body)
    {
        try
        {
            return _answer(target, body);
        }
        catch (Exception ex)
        {
            return $$"""{"error":"the door failed: {{ex.GetType().Name}}"}""";
        }
    }

    /// <summary>
    /// A string argument as VBA actually sends it: a literal arrives as VT_BSTR, a String
    /// variable ByRef as VT_BYREF|VT_BSTR, and a Variant variable as VT_BYREF|VT_VARIANT
    /// wrapping either. Anything else is null, answered as a usage error by the caller above.
    /// </summary>
    private static unsafe string? StringOf(ComVariant* variant)
    {
        if (variant is null)
        {
            return null;
        }

        var kind = variant->VarType;

        if (kind == VarEnum.VT_BSTR)
        {
            var text = variant->GetRawDataRef<nint>();
            return text == 0 ? string.Empty : Marshal.PtrToStringBSTR(text);
        }

        if (kind == (VarEnum.VT_BYREF | VarEnum.VT_BSTR))
        {
            var slot = (nint*)variant->GetRawDataRef<nint>();
            return slot is null ? null : *slot == 0 ? string.Empty : Marshal.PtrToStringBSTR(*slot);
        }

        if (kind == (VarEnum.VT_BYREF | VarEnum.VT_VARIANT))
        {
            var inner = (ComVariant*)variant->GetRawDataRef<nint>();
            return inner is null ? null : StringOf(inner);
        }

        return null;
    }
}
#endif
