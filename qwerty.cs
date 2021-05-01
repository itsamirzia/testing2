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
		string strAccountNumber = string.Empty;
		string strPolicy = string.Empty;
		Application _app = null;
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
				if(strAccountNumber==string.Empty)
				{
					string connectionString = "Server = EPDW; Database = RGHdw; user=dw; password=dw";
					string sqlQuery = "select cusnum as AccountNumber from dbo.EdgCusIns (nolock) where policynum = '"+strPolicy+"'";
					DataTable dt = SelectDataRows(connectionString, sqlQuery);
					if(dt.Rows.Count>0)
					{
						ModifyKeywordInCurrentDocument(doc,"AR - Account Number",dt.Rows[0][0].ToString());
					}
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
		                        strAccountNumber = keyword.IsBlank ? string.Empty : keyword.Value.ToString();
		                        break;
							case "AR - Patient ID":
		                        strPolicy = keyword.IsBlank ? string.Empty : keyword.Value.ToString();
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
        #endregion
    }
}
