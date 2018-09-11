using System;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections;
using System.ComponentModel;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using AWSNotificationMessageFormatter.Constants;

namespace AWSNotificationMessageFormatter
{
    public partial class AWSNotificationMessageFormatter
    {

		Folder inbox;

		private void AWSNotificationMessageFormatter_Startup(object sender, EventArgs e)
        {

			Application.NewMailEx += new ApplicationEvents_11_NewMailExEventHandler(NewMailEx_Handler);


			BackgroundWorker inboxProcessor = new BackgroundWorker();
			inboxProcessor.DoWork += new DoWorkEventHandler(ProcessInbox_Handler);
			inboxProcessor.RunWorkerAsync();
		
		}

		private void AWSNotificationMessageFormatter_Shutdown(object sender, EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that 
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		private void ProcessInbox_Handler(object sender, EventArgs e)
		{
			inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;

			IEnumerable<object> inboxObjects = ((IEnumerable)inbox.Items).Cast<object>();

			IEnumerable<MailItem> inboxItems = inboxObjects.Where(i => i is MailItem).Cast<MailItem>().OrderByDescending(i => i.ReceivedTime).Take(200).ToList();

			foreach (MailItem item in inboxItems)
			{
				ProcessMessage(item);
			}

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
			bool isAwsNotification = 
				mailItem.SenderEmailAddress == AWSNotification.EMAIL_ADDRESS &&
				mailItem.Body != null;

			string newSubject = mailItem.Subject;

			if (isAwsNotification)
			{
				if (mailItem.Body.StartsWith(CodeCommit.PULL_REQUEST_BODY_PREFIX))
				{
					newSubject = GetNewPullRequestSubject(mailItem.Body);
				}
				else if (mailItem.Body.StartsWith(CodePipeline.BODY_PREFIX))
				{
					newSubject = GetNewCodePipelineSubject(mailItem.Body);
				}
				else if (mailItem.Body.StartsWith(CodeCommit.CODE_COMPARE_COMMENT_BODY_PREFIX))
				{
					newSubject = GetNewCodeCompareSubject(mailItem.Body);
				}

				if (newSubject != mailItem.Subject)
				{
					mailItem.Subject = newSubject;
					mailItem.Save();
				}
			}
		}

		private string GetNewPullRequestSubject(string body)
		{
			int startUserIndex = body.IndexOf("arn:aws:") + 8;
			string[] userTokens = body.Substring(startUserIndex).Split(' ')[0].Split('/');
			string user = userTokens[userTokens.Length - 1];

			string action = "updated";
			if (body.Contains(CodeCommit.PULL_REQUEST_COMMENT_IDENTIFIER))
			{
				action = "commented on";
			}
			else if (body.Contains(CodeCommit.PULL_REQUEST_CREATED_IDENTIFIER))
			{
				action = "created";
			}
			else if (body.Contains(CodeCommit.PULL_REQUEST_MERGED_IDENTIFIER))
			{
				action = "merged";
			}

			int pullRequestNumberIndex = -1;
			string pullRequestNumber = "";

			if (body.Contains(CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER))
			{
				pullRequestNumberIndex = body.IndexOf(CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER);
				pullRequestNumber = Regex.Replace(body.Substring(pullRequestNumberIndex + CodeCommit.PULL_REQUEST_NUMBER_IDENTIFIER.Length).Split(' ')[0], "[^0-9]", "");
			}
			else if (body.Contains(CodeCommit.PULL_REQUEST_ALT_NUMBER_IDENTIFIER))
			{
				pullRequestNumberIndex = body.IndexOf(CodeCommit.PULL_REQUEST_ALT_NUMBER_IDENTIFIER);
				pullRequestNumber = Regex.Replace(body.Substring(pullRequestNumberIndex + CodeCommit.PULL_REQUEST_ALT_NUMBER_IDENTIFIER.Length).Split(' ')[0], "[^0-9]", "");
			}

			string repository = body.Substring(CodeCommit.PULL_REQUEST_BODY_PREFIX.Length).Split(' ')[0];
			string reference = $"pull request {pullRequestNumber} in {repository}";

			return $"{user} {action} {reference}";
		}

		private string GetNewCodePipelineSubject(string body)
		{
			string status = "pipeline event";
			if (body.Contains(CodePipeline.FAILED_IDENTIFIER))
			{
				status = "pipeline failed";
			}

			string eventId = body.Substring(body.IndexOf(CodePipeline.EVENT_ID_IDENTIFIER) + CodePipeline.EVENT_ID_IDENTIFIER.Length).Split(' ')[0];

			string repository = CodePipeline.REPOSITORY;

			string[] stageTokens = body.Substring(0, body.IndexOf(CodePipeline.STAGE_IDENTIFIER)).Split(' ');

			string stage = stageTokens[stageTokens.Length - 1];

			return $"{status} for {repository} in {stage.ToLower()} on execution: {eventId}";
		}

		private string GetNewCodeCompareSubject(string body)
		{
			int startUserIndex = body.IndexOf("arn:aws:") + 8;
			string[] userTokens = body.Substring(startUserIndex).Split(' ')[0].Split('/');
			string user = userTokens[userTokens.Length - 1];

			string action = "updated";
			if (body.Contains(CodeCommit.CODE_COMPARE_COMMENT_IDENTIFIER))
			{
				action = "commented on";
			}

			return $"{user} {action} a code comparison";
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
