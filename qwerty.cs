// Skeleton generated by Hyland Unity Editor on 3/17/2021 12:18:40 AM
namespace testam
{
    using System;
    using System.Text;
    using Hyland.Unity;
    using Hyland.Unity.CodeAnalysis;
    using Hyland.Unity.Workflow;
	using System.IO;
	using System.Data;
	using System.Data.SqlClient;
    
    
    /// <summary>
    /// testam
    /// </summary>
    public class testam2 : Hyland.Unity.IWorkflowScript
    {

		Application _app = null;
		string sAccount = string.Empty;
		string sPatientAuthSignedDt = string.Empty;
		string sPatientAuthSignedUser = string.Empty;
		string sPatientAuthToRGH = string.Empty;
		string sPatientAuthMMDDYYYY = string.Empty;
		bool foundAccount = false;
		bool foundPatientAuthSignedDt = false;
		bool foundPatientAuthSignedUser = false;
		bool foundPatientAuthToRGH = false;
		
        #region IWorkflowScript
        /// <summary>
        /// Implementation of <see cref="IWorkflowScript.OnWorkflowScriptExecute" />.
        /// <seealso cref="IWorkflowScript" />
        /// </summary>
        /// <param name="app"></param>
        /// <param name="args"></param>
        public void OnWorkflowScriptExecute(Hyland.Unity.Application app, Hyland.Unity.WorkflowEventArgs args)
        {
            
			try
			{
				_app = app;
				Document doc = args.Document;
				SetKeywordValues(doc);
				if(foundAccount && foundPatientAuthSignedDt && foundPatientAuthSignedUser)
				{
					string sTimeStamp = System.DateTime.Now.ToString("MMddyyyyHHmmss");
					string oFilePath = @"Cardinalhealthy.net\applications\EC500\ONB\PRD_SRC\BXOEXP\MR\PatientAuth\";
					string oFileName = "PatientAuth"+"_"+sAccount+"_"+sTimeStamp+".txt";
					string oFilePathName = oFilePath+oFileName;
					string oPatientInfo = sAccount+"|"+sPatientAuthMMDDYYYY+"|"+sPatientAuthSignedUser;
					File.AppendAllText(oFilePathName, oPatientInfo);
					ModifyKeywordInCurrentDocument(doc, "MR - Patient Auth RGH File",oFilePathName);
				}
				
			}
			catch(Exception ex)
			{
				app.Diagnostics.Write(ex);
			}
			
			
        }
		
		private void ModifyKeywordInCurrentDocument(Document doc, string keywordType, string keywordValue)
		{
			using(DocumentLock documentLock = doc.LockDocument())
			{
				if(documentLock.Status == DocumentLockStatus.LockObtained)
				{
					KeywordModifier keymod = doc.CreateKeywordModifier();
					KeywordType keyType = _app.Core.KeywordTypes.Find(keywordType);
					if(keyType == null)
					{
						keymod.AddKeyword(keywordType,keywordValue);		
						
					}
					else
					{
						KeywordRecord keyRec = doc.KeywordRecords.Find(keyType);
						KeywordRecordType keyRecType = keyRec.KeywordRecordType;			
						
						Keyword newKeyword = keyType.CreateKeyword(keywordValue);
						
						if(keyRecType.RecordType== RecordType.MultiInstance)
						{
							EditableKeywordRecord editKeyRec = keyRec.CreateEditableKeywordRecord();
							Keyword keyword = editKeyRec.Keywords.Find(keywordType);
							if(keyword != null)
							{
								editKeyRec.UpdateKeyword(keyword, newKeyword);
							}
							else
							{
								editKeyRec.AddKeyword(keywordType, keywordValue);
							}
							
							keymod.AddKeywordRecord(editKeyRec);
							
						}
						else
						{
							Keyword keyword = keyRec.Keywords.Find(keywordType);
							if(keyword!=null)
								keymod.UpdateKeyword(keyword, newKeyword);
							else
								keymod.AddKeyword(keywordType, keywordValue);
						}
					}
					keymod.ApplyChanges();					
				}
			}
		}
		private void SetKeywordValues(Document document)
		{
			try
		   	{
		   		foreach(KeywordRecord keywordRecord in document.KeywordRecords)
		        {
		        	foreach(Keyword keyword in keywordRecord.Keywords)
		            {
		            	switch (keyword.KeywordType.Name)
		                {
							case "Account # A":
								sAccount = keyword.IsBlank? string.Empty: keyword.Value.ToString().Trim();
								foundAccount=true;
								break;
							case "MR - Patient Auth Signed Date":
								sPatientAuthSignedDt = keyword.IsBlank? string.Empty: keyword.Value.ToString().Trim();
								string sPatientAuthSignedDtOnly = sPatientAuthSignedDt.Substring(0,Math.Min(10,sPatientAuthSignedDt.Length));
								string[] sSplit = sPatientAuthSignedDtOnly.Split('-');
								sPatientAuthMMDDYYYY = sSplit[1] +"/" + sSplit[2] + "/" + sSplit[0];
								foundPatientAuthSignedDt = true;
								break;
							case "MR - Patient Auth Signed User":
								sPatientAuthSignedUser = keyword.IsBlank? string.Empty: keyword.Value.ToString().Trim();
								foundPatientAuthSignedUser = true;
								break;
							case "MR - Patient Auth RGH File":
								sPatientAuthToRGH = keyword.IsBlank? string.Empty: keyword.Value.ToString().Trim();
								foundPatientAuthToRGH = true;
								break;
		                    
						}
					}
				}
		   }
		   catch(Exception ex)
		   {
		   		_app.Diagnostics.Write(ex);
		   }
		}
		
        #endregion
    }
}
