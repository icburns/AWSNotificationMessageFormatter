namespace AWSNotificationMessageFormatter.Constants
{
	class CodeCommit
	{
		public static string BODY_PREFIX => "\"A pull request event occurred in the following AWS CodeCommit repository: ";
		public static string COMMENT_IDENTIFIER => "made a comment or replied to a comment. ";
		public static string CREATED_IDENTIFIER => "made the following PullRequest ";
		public static string MERGED_IDENTIFIER => "The status is merged. ";
		public static string PULL_REQUEST_NUMBER_IDENTIFIER => "the following Pull Request: ";
		public static string ALT_PULL_REQUEST_NUMBER_IDENTIFIER => "the following PullRequest ";

	}
}
