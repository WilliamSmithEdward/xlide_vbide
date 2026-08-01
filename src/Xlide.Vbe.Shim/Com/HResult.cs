namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// The HRESULT values this server returns or inspects. Named so that call sites read as intent
/// rather than as hexadecimal.
/// </summary>
internal static class HResult
{
    public const int Ok = 0;
    public const int False = 1;

    public const int NotImplemented = unchecked((int)0x80004001);
    public const int NoInterface = unchecked((int)0x80004002);
    public const int Fail = unchecked((int)0x80004005);
    public const int InvalidArg = unchecked((int)0x80070057);
    public const int OutOfMemory = unchecked((int)0x8007000E);
    public const int Unexpected = unchecked((int)0x8000FFFF);

    public const int ClassNotAvailable = unchecked((int)0x80040111);
    public const int ClassNotRegistered = unchecked((int)0x80040154);
    public const int NoAggregation = unchecked((int)0x80040110);

    public const int DispMemberNotFound = unchecked((int)0x80020003);
    public const int DispBadParamCount = unchecked((int)0x8002000E);
    public const int DispUnknownName = unchecked((int)0x80020006);

    /// <summary>OLE_S_USEREG. Tells the container to read the answer from the class registration.</summary>
    public const int OleUseRegistry = 0x00040000;

    /// <summary>OLEOBJ_S_INVALIDVERB. A success code: the verb was understood and refused.</summary>
    public const int OleInvalidVerb = 0x00040180;

    public const int OleAdviseNotSupported = unchecked((int)0x80040003);
    public const int OleNoConnection = unchecked((int)0x80040004);
}
