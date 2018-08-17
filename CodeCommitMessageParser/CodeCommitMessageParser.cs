using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace CodeCommitMessageParser
{
    public partial class CodeCommitMessageParser
    {
		const string CODE_COMMIT_MESSAGE_SUBJECT = "AWS Notification Message";
		const string CODE_COMMIT_EMAIL = "no-reply@sns.amazonaws.com";
		const string CODE_COMMIT_BODY_PREFIX = "\"A pull request event occurred in the following AWS CodeCommit repository: ";
		const string COMMENT_IDENTIFIER = "made a comment or replied to a comment ";
		const string CREATED_IDENTIFIER = "made the following PullRequest ";
		const string MERGED_IDENTIFIER = "The status is merged. ";
		const string PULL_REQUEST_NUMBER_IDENTIFIER = "the following PullRequest";

		bool fullParseRequired = true;

		Explorer currentExplorer;

		private void CodeCommitMessageParser_Startup(object sender, EventArgs e)
        {
			currentExplorer = Application.ActiveExplorer();
			currentExplorer.SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Handler);
		}

		private void CodeCommitMessageParser_Shutdown(object sender, EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that 
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		private void CurrentExplorer_Handler()
		{
			if (!fullParseRequired)
			{
				return;
			}

			Folder selectedFolder =
				this.Application.ActiveExplorer().CurrentFolder as Folder;

			if (selectedFolder.Name != "Inbox")
			{
				return;
			}

			selectedFolder.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(ItemAdd_Handler);

			foreach (var item in selectedFolder.Items) {

				if (item is MailItem)
				{
					MailItem mailItem = item as MailItem;
					if (mailItem.Subject == CODE_COMMIT_MESSAGE_SUBJECT &&
						mailItem.SenderEmailAddress == CODE_COMMIT_EMAIL &&
						mailItem.Body != null &&
						mailItem.Body.StartsWith(CODE_COMMIT_BODY_PREFIX))
					{
						mailItem.Subject = GetNewSubject(mailItem.Body);
						mailItem.Save();
					}
				}
			}

			fullParseRequired = false;
		}

		private void ItemAdd_Handler(object newItem)
		{
			if (newItem is MailItem)
			{
				MailItem mailItem = newItem as MailItem;
				if (mailItem.Subject == CODE_COMMIT_MESSAGE_SUBJECT &&
					mailItem.SenderEmailAddress == CODE_COMMIT_EMAIL &&
					mailItem.Body != null &&
					mailItem.Body.StartsWith(CODE_COMMIT_BODY_PREFIX))
				{
					mailItem.Subject = GetNewSubject(mailItem.Body);
					mailItem.Save();
				}
			}
		}

		private string GetNewSubject(string body)
		{
			int startUserIndex = body.IndexOf("arn:aws:") + 8;
			string[] userTokens = body.Substring(startUserIndex).Split(' ')[0].Split('/');
			string user = userTokens[userTokens.Length - 1];

			string action = "";
			if (body.Contains(COMMENT_IDENTIFIER))
			{
				action = "commented on";
			}
			else if (body.Contains(CREATED_IDENTIFIER))
			{
				action = "created";
			}
			else if (body.Contains(MERGED_IDENTIFIER))
			{
				action = "merged";
			}
			else
			{
				action = "updated";
			}

			int pullRequestNumberIndex = body.IndexOf(PULL_REQUEST_NUMBER_IDENTIFIER + ": ") != -1 ?
											body.IndexOf(PULL_REQUEST_NUMBER_IDENTIFIER + ": ") :
											body.IndexOf(PULL_REQUEST_NUMBER_IDENTIFIER + " ");
			string pullRequestNumber = Regex.Replace(body.Substring(pullRequestNumberIndex + PULL_REQUEST_NUMBER_IDENTIFIER.Length).Split(' ')[1], "[^0-9]", "");
			string repository = body.Substring(CODE_COMMIT_BODY_PREFIX.Length).Split(' ')[0];
			string reference = $"pull request {pullRequestNumber} in {repository}";

			return $"{user} {action} {reference}";
		}

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
        {
            Startup += new EventHandler(CodeCommitMessageParser_Startup);
            Shutdown += new EventHandler(CodeCommitMessageParser_Shutdown);
        }
        
    }
}
