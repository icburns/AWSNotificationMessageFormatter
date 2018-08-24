using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using AWSNotificationMessageFormatter.Constants;

namespace AWSNotificationMessageFormatter
{
    public partial class AWSNotificationMessageFormatter
    {

		Folder inbox;

		private void AWSNotificationMessageFormatter_Startup(object sender, EventArgs e)
        {
			inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;

			Application.NewMailEx += new ApplicationEvents_11_NewMailExEventHandler(NewMailEx_Handler);

			//newMailEx seems to be more reliable
			//inbox.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(ItemAdd_Handler);

			foreach (var item in inbox.Items)
			{
				if (item is MailItem)
				{
					MailItem mailItem = item as MailItem;

					ProcessMessage(mailItem);
				}
			}
			
		}

		private void AWSNotificationMessageFormatter_Shutdown(object sender, EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that 
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		private void NewMailEx_Handler(string newItemId)
		{
			var mapiInbox = inbox as MAPIFolder;
			object newItem = Application.GetNamespace("MAPI").GetItemFromID(newItemId, mapiInbox.StoreID);
			MailItem mailItem;

			if (newItem is MailItem)
			{
				mailItem = newItem as MailItem;

				ProcessMessage(mailItem);
			}
		}

		private void ProcessMessage(MailItem mailItem)
		{
			bool isAwsNotification = mailItem.Subject == AWSNotification.MESSAGE_SUBJECT &&
				mailItem.SenderEmailAddress == AWSNotification.EMAIL_ADDRESS &&
				mailItem.Body != null;

			if (isAwsNotification)
			{
				if (mailItem.Body.StartsWith(CodeCommit.BODY_PREFIX))
				{
					mailItem.Subject = GetNewCodeCommitSubject(mailItem.Body);
				}
				else if (mailItem.Body.StartsWith(CodePipeline.BODY_PREFIX))
				{
					mailItem.Subject = GetNewCodePipelineSubject(mailItem.Body);
				}

				mailItem.Save();
			}
		}

		private string GetNewCodeCommitSubject(string body)
		{
			int startUserIndex = body.IndexOf("arn:aws:") + 8;
			string[] userTokens = body.Substring(startUserIndex).Split(' ')[0].Split('/');
			string user = userTokens[userTokens.Length - 1];

			string action = "";
			if (body.Contains(CodeCommit.COMMENT_IDENTIFIER))
			{
				action = "commented on";
			}
			else if (body.Contains(CodeCommit.CREATED_IDENTIFIER))
			{
				action = "created";
			}
			else if (body.Contains(CodeCommit.MERGED_IDENTIFIER))
			{
				action = "merged";
			}
			else
			{
				action = "updated";
			}

			int pullRequestNumberIndex = body.IndexOf(CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER + ": ") != -1 ?
											body.IndexOf(CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER + ": ") :
											body.IndexOf(CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER + " ");
			string pullRequestNumber = Regex.Replace(body.Substring(pullRequestNumberIndex + CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER.Length).Split(' ')[1], "[^0-9]", "");
			string repository = body.Substring(CodeCommit.BODY_PREFIX.Length).Split(' ')[0];
			string reference = $"pull request {pullRequestNumber} in {repository}";

			return $"{user} {action} {reference}";
		}

		private string GetNewCodePipelineSubject(string body)
		{
			string status = "";
			if (body.Contains(CodePipeline.FAILED_IDENTIFIER))
			{
				status = "pipeline failed";
			}
			else
			{
				status = "pipeline event";
			}

			string eventId = body.Substring(body.IndexOf(CodePipeline.EVENT_ID_IDENTIFIER) + CodePipeline.EVENT_ID_IDENTIFIER.Length).Split(' ')[0];

			string repository = CodePipeline.REPOSITORY;

			string[] stageTokens = body.Substring(0, body.IndexOf(CodePipeline.STAGE_IDENTIFIER)).Split(' ');

			string stage = stageTokens[stageTokens.Length - 1];

			return $"{status} for {repository} in {stage.ToLower()} on execution: {eventId}";
		}



		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
        {
            Startup += new EventHandler(AWSNotificationMessageFormatter_Startup);
            Shutdown += new EventHandler(AWSNotificationMessageFormatter_Shutdown);
        }
        
    }
}
