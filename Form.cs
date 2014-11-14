using Aspose.BarCode;
using FISCA.Data;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words.Drawing;
using Aspose.Words;
using System.IO;
using SmartSchool.ePaper;
using Aspose.Words.Reporting;

namespace InvitationParents
{
	public partial class Form : BaseForm
	{
		/// <summary>
		/// 學生電子報表
		/// </summary>
		SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }


		Document _template = new Document();
		BackgroundWorker bgw = new BackgroundWorker();
		private List<string> studentIds;
		private QueryHelper queryHelper;
		private string _type;
		MemoryStream studentTemplate = new MemoryStream(Properties.Resources.Student_QRcode);

		public Form(List<string> selectStudentIds, string type)
		{
			InitializeComponent();
			studentIds = selectStudentIds;
			this.queryHelper = new QueryHelper();


			bgw.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
			bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);

			_type = type;
		}

		private void buttonX1_Click(object sender, EventArgs e)
		{
			if (_type == "student")
			{
				_template = new Document(new MemoryStream(Properties.Resources.Student_QRcode));
				bgw.RunWorkerAsync(studentIds);
			}
			else if (_type == "class")
			{
				_template = new Document(new MemoryStream(Properties.Resources.Class_QRcode));
				bgw.RunWorkerAsync(studentIds);
			}
		}

		private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			List<string> student = (List<string>)e.Argument;

			string DSNSName = FISCA.Authentication.DSAServices.AccessPoint;
			string schoolName = K12.Data.School.ChineseName;
			string ids = string.Join(",", student);

			string sql = "select student.id,student.parent_code,seat_no,name,class.grade_year,class.class_name from student";
			sql += " join class on class.id = student.ref_class_id where student.status in (1,2) and student.id in (" + ids + ") order by class.grade_year,class.display_order,class.class_name,student.seat_no";
			DataTable dt = queryHelper.Select(sql); ;
			Document doc = new Document();
			Dictionary<string, object> merge = new Dictionary<string, object>();
			int count = 0, pageSize = 9;
			int index = 0;

			string previous_cn = string.Empty;

			//start
			if (_type == "student")
			{
				foreach (DataRow row in dt.Rows)
				{
					string gradeYear = row["grade_year"] + "",
						studentName = row["name"] + "",
						seatNo = row["seat_no"] + "",
						className = row["class_name"] + "",
						studentId = row["id"] + "",
						parentsCode = row["parent_code"] + "";

					Document perPage = _template.Clone();
					merge.Clear();
					merge.Add("年級", gradeYear);
					merge.Add("班級名稱", className);
					merge.Add("座號", seatNo);
					merge.Add("學生姓名", studentName);
					merge.Add("學校名稱", schoolName);
					merge.Add("QRCODE", parentsCode + "@" + DSNSName);
					merge.Add("家長代碼", parentsCode);
					perPage.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
					perPage.MailMerge.Execute(merge.Keys.ToArray<string>(), merge.Values.ToArray<object>());
					perPage.MailMerge.DeleteFields();

					doc.Sections.Add(doc.ImportNode(perPage.Sections[0], true));
				}
			}

			if (_type == "class")
			{
				Dictionary<string, Dictionary<string, string>> classes = new Dictionary<string, Dictionary<string, string>>();

				int offset = 0, size = 9;
				foreach (DataRow row in dt.Rows)
				{
					string gradeYear = row["grade_year"] + "",
						studentName = row["name"] + "",
						seatNo = row["seat_no"] + "",
						className = row["class_name"] + "",
						studentId = row["id"] + "",
						parentsCode = row["parent_code"] + "";

					if (!classes.ContainsKey(className))
					{
						classes.Add(className, new Dictionary<string, string>());
						offset = 0;
					}

					if (!classes[className].ContainsKey("班級名稱"))
						classes[className].Add("班級名稱", className);

					if (!classes[className].ContainsKey("學校名稱"))
						classes[className].Add("學校名稱", className);

					classes[className].Add("年級" + offset % size, gradeYear);
					classes[className].Add("座號" + offset % size, seatNo);
					classes[className].Add("學生姓名" + offset % size, studentName);
					classes[className].Add("QRCODE" + offset % size, parentsCode + "@" + DSNSName);
					classes[className].Add("家長代碼" + offset % size, parentsCode);

					if (offset % size == 0)
					{
						Document perPage = _template.Clone();
						perPage.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
						perPage.MailMerge.Execute(merge.Keys.ToArray<string>(), merge.Values.ToArray<object>());
						perPage.MailMerge.DeleteFields();

						doc.Sections.Add(doc.ImportNode(perPage.Sections[0], true));
					}

					offset++;
				}


				//if (!merge.ContainsKey("班級名稱"))
				//	merge.Add("班級名稱", className);

				//if (!merge.ContainsKey("學校名稱"))
				//	merge.Add("學校名稱", schoolName);

				//merge.Add("年級" + index, gradeYear);
				//merge.Add("座號" + index, seatNo);
				//merge.Add("學生姓名" + index, studentName);
				//merge.Add("QRCODE" + index, parentsCode + "@" + DSNSName);
				//merge.Add("家長代碼" + index, parentsCode);

				//if ((index == (pageSize - 1)) || (className != previous_cn && count > 0))
				//{
				//	Document perPage = _template.Clone();
				//	perPage.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
				//	perPage.MailMerge.Execute(merge.Keys.ToArray<string>(), merge.Values.ToArray<object>());
				//	perPage.MailMerge.DeleteFields();

				//	doc.Sections.Add(doc.ImportNode(perPage.Sections[0], true));
				//	merge.Clear();
				//}

				//count++;
				//previous_cn = className;
				//index = (count % pageSize);

			}
			//end

			doc.Sections.RemoveAt(0);
			e.Result = doc;
		}
		class MailMerge_MergeField : IFieldMergingCallback
		{

			private List<string> qrcodeFieldNames = new List<string>()
			 {
				"QRCODE",
				"QRCODE0",
				"QRCODE1",
				"QRCODE2",
				"QRCODE3",
				"QRCODE4",
				"QRCODE5",
				"QRCODE6",
				"QRCODE7",
				"QRCODE8",
				"QRCODE9",
				"QRCODE10",
				"QRCODE11",
				"QRCODE12"
			 };
			public void FieldMerging(FieldMergingArgs args)
			{
				FieldMergingArgs e = args;
				if (!qrcodeFieldNames.Contains(e.DocumentFieldName))
					return;
				DocumentBuilder builder = new DocumentBuilder(e.Document);
				builder.MoveToField(e.Field, true);
				e.Field.Remove();
				if (e.FieldValue != null && e.FieldValue.ToString() != "")
				{
					BarCodeBuilder bb = new BarCodeBuilder(e.FieldValue.ToString(), Symbology.QR);
					bb.GraphicsUnit = GraphicsUnit.Millimeter;
					bb.AutoSize = false;
					//QRcode size
					bb.xDimension = 1.2f;
					bb.yDimension = 1.2f;
					//Image size
					bb.ImageWidth = 42f;
					bb.ImageHeight = 31.5f;
					MemoryStream stream = new MemoryStream();
					bb.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
					builder.InsertImage(stream);
				}
			}
			public void ImageFieldMerging(Aspose.Words.Reporting.ImageFieldMergingArgs args)
			{
			}
		}

		void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			if (e.Error != null)
			{
				FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤:" + e.Error.Message);
				return;
			}

			Document document = (Document)e.Result;
			string inputReportName1 = "學生家長二維條碼邀請函";
			string inputReportName2 = "班級家長二維條碼邀請函";
			string reportName = "";
			if (_type == "student")
			{
				reportName = inputReportName1;
			}
			if (_type == "class")
			{
				reportName = inputReportName2;
			}

			System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
			sd.Title = "另存新檔";
			sd.FileName = reportName + ".doc";
			sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
			if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				try
				{
					document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
					System.Diagnostics.Process.Start(sd.FileName);
				}
				catch (Exception ex)
				{
					FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
					return;
				}
			}
		}

		private void buttonX2_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
		{
			if (checkBoxX1.ThreeState == true)
			{
				Document doc = new Document();
				doc.Sections.Clear();

				#region 班級學生的電子報表
				//建立一個學生電子報表
				//傳入參數 : 報表名稱,學年度,學期,類型(學生/班級/教師/課程)
				paperForStudent = new SmartSchool.ePaper.ElectronicPaper("學生家長代碼邀請函 for App", School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

				//學生個人的文件
				Document each_page = new Document(studentTemplate);
				MemoryStream stream = new MemoryStream();
				each_page.Save(stream, SaveFormat.Doc);
				doc.Sections.Add(doc.ImportNode(each_page.Sections[0], true)); //合併至doc

				List<string> ClassID = K12.Presentation.NLDPanels.Class.SelectedSource; //取得畫面上所選班級的ID清單
				List<StudentRecord> srList = Student.SelectByClassIDs(ClassID); //依據班級ID,取得學生物件
				foreach (StudentRecord sr in srList)
				{
					//傳參數給PaperItem
					//格式 / 內容 / 對象的系統編號
					paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, sr.ID));
				}

				//開始上傳
				SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForStudent);
				#endregion
			}
		}

		private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
			sd.Title = "另存新檔";
			sd.FileName = "合併欄位.doc";
			sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
			if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				try
				{
					Document document = new Document(new MemoryStream(Properties.Resources.QRCODE_合併欄位));
					document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
					System.Diagnostics.Process.Start(sd.FileName);
				}
				catch (Exception ex)
				{
					FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
					return;
				}
			}
		}
	}
}
