
namespace Office.Shared
{
	public delegate void LogDelegate(string message);
	class Logger
	{
		protected LogDelegate loggers;
		protected static Logger logger=new Logger();
		protected Logger()
		{
			loggers = null;
		}

		public static void addFunc(LogDelegate log){
			logger.loggers=log;
		}

		public static void log(string message) {
			if (logger.loggers != null) {
				logger.loggers(message);
			}
		}
	}
}
