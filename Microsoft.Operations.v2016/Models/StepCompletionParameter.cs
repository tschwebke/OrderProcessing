/// <summary>
/// This class is typically used with any background worker threads, to report progress or result
/// from a working thread. example of usage: StepCompletionParameter param = (StepCompletionParameter)e.Result;
/// </summary>
public class StepCompletionParameter
{
    public StepCompletionParameter(int errorCode, string anythingToSay)
    {
        ErrorCode = errorCode;
        ErrorMessage = anythingToSay;
    }

    public int ErrorCode { get; set; }
    public string ErrorMessage { get; set; }
}