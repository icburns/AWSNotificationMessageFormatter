namespace AWSNotificationMessageFormatter.Constants
{
	class CodeCommit
	{
		public static string PULL_REQUEST_BODY_PREFIX => "\"A pull request event occurred in the following AWS CodeCommit repository: ";
		public static string PULL_REQUEST_COMMENT_IDENTIFIER => "made a comment or replied to a comment. ";
		public static string PULL_REQUEST_CREATED_IDENTIFIER => "made the following PullRequest ";
		public static string PULL_REQUEST_MERGED_IDENTIFIER => "The status is merged. ";
		public static string PULL_REQUEST_NUMBER_IDENTIFIER => "the following Pull Request: ";
		public static string PULL_REQUEST_ALT_NUMBER_IDENTIFIER => "the following PullRequest ";

		public static string CODE_COMPARE_COMMENT_BODY_PREFIX => "\"A comment event occurred in the following AWS CodeCommit repository: ";
		public static string CODE_COMPARE_COMMENT_IDENTIFIER => "made a comment or replied to a comment. ";

	}
}
