namespace AWSNotificationMessageFormatter.Constants
{
	class CodePipeline
	{
		public static string BODY_PREFIX => "\"The edfred-cicd-datalake-wind pipeline ";
		public static string FAILED_IDENTIFIER => "failed during execution with ID ";
		//TODO: ianb - this should be identified from the message after rework - 20180824
		public static string REPOSITORY => "edfred-cicd-datalake-wind";
		public static string STAGE_IDENTIFIER => " stage\"";
		public static string EVENT_ID_IDENTIFIER => "with ID ";
	}
}
