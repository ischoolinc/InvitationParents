using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.Report;
using Aspose.Words;
using System.IO;

namespace InvitationParents
{
    public class ReportConfig
    {
        public  ReportConfiguration _RptConfig;
        public ReportConfig()
        {
            _RptConfig = new ReportConfiguration("InvitationParents_Config");
        }

        /// <summary>
        /// class,student
        /// </summary>
        /// <param name="type"></param>
        public  Document LoadTemplate(string type)
        {            
            Document value=null;         
            if (type == "class")
            {
                try
                {
                    string strClass = _RptConfig.GetString("TemplateClass", "");
                    if (string.IsNullOrEmpty(strClass))
                    {
                        // 使用預設                    
                        value = new Document(new MemoryStream(Properties.Resources.Class_QRcode));
                        MemoryStream stream = new MemoryStream ();
                        value.Save(stream,Aspose.Words.SaveFormat.Doc);
                        // 回存
                        
                        _RptConfig.SetString("TemplateClass", Convert.ToBase64String(stream.ToArray()));
                        _RptConfig.Save();

                    }
                    else
                    {
                        value = new Document(new MemoryStream(Convert.FromBase64String(strClass)));
                    }
                }
                catch(Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("載入班級範本失敗,", ex.Message);
                }
            }

            if(type =="student")
            {
                try
                {
                    string strStudent = _RptConfig.GetString("TemplateStudent", "");
                    if (string.IsNullOrEmpty(strStudent))
                    {
                        // 使用預設                    
                        value = new Document(new MemoryStream(Properties.Resources.學生家長二維條碼邀請函201609));
                        
                        // 回存                        
                        MemoryStream stream = new MemoryStream();
                        value.Save(stream, Aspose.Words.SaveFormat.Doc);                       

                        _RptConfig.SetString("TemplateStudent", Convert.ToBase64String(stream.ToArray()));
                        _RptConfig.Save();

                    }
                    else
                    {
                        value = new Document(new MemoryStream(Convert.FromBase64String(strStudent)));
                    }                

                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("載入學生範本失敗,", ex.Message);
                }

            }

            return value;
        }

        /// <summary>
        /// class,student
        /// </summary>
        /// <param name="DocTemp"></param>
        /// <param name="type"></param>
        public void SaveTemplate(Document DocTemp,string type)
        {            
            if(type=="class")
            {
                try
                {
                    MemoryStream ms = new MemoryStream();
                    DocTemp.Save(ms, Aspose.Words.SaveFormat.Doc);                    
                    _RptConfig.SetString("TemplateClass", Convert.ToBase64String(ms.ToArray()));
                    _RptConfig.Save();

                }catch(Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存班級範本失敗,", ex.Message);
                }
            }

            if(type=="student")
            {
                try
                {
                    MemoryStream ms = new MemoryStream();
                    DocTemp.Save(ms, Aspose.Words.SaveFormat.Doc);                    
                    _RptConfig.SetString("TemplateStudent", Convert.ToBase64String(ms.ToArray()));
                    _RptConfig.Save();

                }catch(Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存學生範本失敗,", ex.Message);
                }
            }
        
        }

    }
}
