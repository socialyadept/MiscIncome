using QBFC16Lib;

namespace QB_MiscIncome_Lib
{
    public class QuickBooksSession : IDisposable
    {
        private QBSessionManager _sessionManager;
        private bool _sessionBegun;
        private bool _connectionOpen;

        public QuickBooksSession(string appName)
        {
            _sessionManager = new QBSessionManager();
            _sessionManager.OpenConnection("", appName);
            _connectionOpen = true;
            _sessionManager.BeginSession("", ENOpenMode.omDontCare);
            _sessionBegun = true;
        }

        public IMsgSetRequest CreateRequestSet()
        {
            IMsgSetRequest requestMsgSet = _sessionManager.CreateMsgSetRequest("US", 16, 0);
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            return requestMsgSet;
        }

        public IMsgSetResponse SendRequest(IMsgSetRequest requestMsgSet)
        {
            return _sessionManager.DoRequests(requestMsgSet);
        }

        public void Dispose()
        {
            if (_sessionBegun)
            {
                _sessionManager.EndSession();
                _sessionBegun = false;
            }
            if (_connectionOpen)
            {
                _sessionManager.CloseConnection();
                _connectionOpen = false;
            }
        }
    }
}
