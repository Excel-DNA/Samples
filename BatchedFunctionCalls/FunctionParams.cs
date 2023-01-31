namespace BatchedFunctionCalls
{
    internal class FunctionParams
    {
        public string Ticker;
        public int Year;

        public readonly TaskCompletionSource<object> result = new();

        public override string ToString() => $"FunctionParams {Ticker}-{Year}";
    }
}
