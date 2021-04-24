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
    public class testam : Hyland.Unity.IWorkflowScript
    {
		string accountName = string.Empty;
		string accountNumber = string.Empty;
		string dateOfBirth = string.Empty;
		Hyland.Unity.Application _app = null;
        #region IWorkflowScript
        /// <summary>
        /// Implementation of <see cref="IWorkflowScript.OnWorkflowScriptExecute" />.
        /// <seealso cref="IWorkflowScript" />
        /// </summary>
        /// <param name="app"></param>
        /// <param name="args"></param>
        public void OnWorkflowScriptExecute(Hyland.Unity.Application app, Hyland.Unity.WorkflowEventArgs args)
        {
			_app = app;
			Document doc = args.Document;
			SetKeywordValues(doc);
			string sName = accountName.Replace("'","");
			string month = dateOfBirth.Substring(0,2);
			string day  = dateOfBirth.Substring(3,2);
			string year  = dateOfBirth.Substring(6,4);
			string sBirthDate = year+"-"+month+"-"+day;
			if(sName != string.Empty && sBirthDate != "" && accountNumber == "")
			{
				string connectionString = "Server = WPEC5009onbsq01.cardinalhealth; Database = ONBASE; user=hsi; password=wstinol";
				string queryString = "select ks101,ks102 from hsi.keysetdata112 (NOLOCK) where ks109 = '"+sBirthDate+"' and left(ks102,len(ks102)- len(substring(ks102,charindex(',',ks102),len(ks102)))) = '"+sName+"'";
				DataTable dt = SelectDataRows(connectionString, queryString);
				foreach(DataRow dr in dt.Rows)
				{
					string sAccount = dr[0].ToString();
					string sAccountName = dr[1].ToString();
					if(sAccount != "" && sAccountName != "")
					{
						ModifyKeywordInCurrentDocument(doc, "AR - Account Number",sAccount);
					}
				}
			}
			
        }
		private DataTable SelectDataRows(string dbConnectionString,string queryString)
		{
			DataTable dt = new DataTable();
			try
			{
				using (SqlConnection connection =  new SqlConnection(dbConnectionString))
			    {
			        SqlDataAdapter adapter = new SqlDataAdapter();
			        adapter.SelectCommand = new SqlCommand(queryString, connection);
			        adapter.Fill(dt);
			    }
			}
			catch(Exception ex)
			{
				_app.Diagnostics.Write("Error while fetching records Method: SelectDataRows ");
				_app.Diagnostics.Write(ex);
			}
			return dt;
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
							editKeyRec.UpdateKeyword(keyword, newKeyword);
							keymod.AddKeywordRecord(editKeyRec);
							
						}
						else
						{
							Keyword keyword = keyRec.Keywords.Find(keywordType);
							keymod.UpdateKeyword(keyword, newKeyword);
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
		                   case "AR - Account Number":
								accountNumber = keyword.IsBlank?string.Empty:keyword.Value.ToString();
								break;
							case "AR - Date of Birth":
								dateOfBirth = keyword.IsBlank?string.Empty:keyword.Value.ToString();
								break;
							case "AR - Account Name":
								accountName = keyword.IsBlank?string.Empty:keyword.Value.ToString();
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

