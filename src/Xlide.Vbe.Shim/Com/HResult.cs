namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// The HRESULT values this server returns or inspects. Named so that call sites read as intent
/// rather than as hexadecimal.
///
/// EVERY ONE OF THESE HAS A CALLER. Ten more used to sit here, four of them the vocabulary of an
/// in-place-activated embedded object - OLE_S_USEREG, OLEOBJ_S_INVALIDVERB, OLE_E_ADVISENOTSUPPORTED
/// and OLE_E_NOCONNECTION - which this add-in has never been. A table of codes nobody returns reads
/// as a contract the server implements, and sends the next reader looking for the paths that would
/// return them.
/// </summary>
internal static class HResult
{
    public const int Ok = 0;
    public const int False = 1;

    public const int Fail = unchecked((int)0x80004005);
    public const int InvalidArg = unchecked((int)0x80070057);

    public const int ClassNotAvailable = unchecked((int)0x80040111);
    public const int NoAggregation = unchecked((int)0x80040110);

    public const int DispMemberNotFound = unchecked((int)0x80020003);
    public const int DispUnknownName = unchecked((int)0x80020006);
}
