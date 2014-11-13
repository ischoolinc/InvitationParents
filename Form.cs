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

namespace InvitationParents
{
	public partial class Form : BaseForm
	{
		/// <summary>
		/// 學生電子報表
		/// </summary>
		SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }

		Document _doc = new Document();
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
				_template = new Document(new MemoryStream(Properties.Resources.Class_QRcode1));
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
			sql += " join class on class.id = student.ref_class_id where student.id in (" + ids + ") order by class.grade_year,class.display_order,class.class_name,student.seat_no";
			DataTable dt = queryHelper.Select(sql); ;
			Dictionary<string, object> merge = new Dictionary<string, object>();
			int count = 0;

			string previous_cn = string.Empty;
			foreach (DataRow row in dt.Rows)
			{
				string gradeYear = row["grade_year"] + "",
					studentName = row["name"] + "",
					seatNo = row["seat_no"] + "",
					className = row["class_name"] + "",
					studentId = row["id"] + "",
					parentsCode = row["parent_code"] + "";
				//start
				if (_type == "student")
				{
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

					_doc.Sections.Add(_doc.ImportNode(perPage.Sections[0], true));
				}

				if (_type == "class")
				{
					if (merge.ContainsKey(className))
						merge.Add("班級名稱", className);

					if (merge.ContainsKey(schoolName))
						merge.Add("學校名稱", schoolName);

					merge.Add("年級" + count, gradeYear);
					merge.Add("座號" + count, seatNo);
					merge.Add("學生姓名" + count, studentName);
					merge.Add("QRCODE" + count, parentsCode + "@" + DSNSName);
					merge.Add("家長代碼" + count, parentsCode);

					count++;

					if (count == 3)
					{
						Document perPage = _template.Clone();

						//perPage.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
						perPage.MailMerge.Execute(merge.Keys.ToArray<string>(), merge.Values.ToArray<object>());
						perPage.MailMerge.DeleteFields();

						_doc.Sections.Add(_doc.ImportNode(perPage.Sections[0], true));

						count = 0;
						merge.Clear();
					}

					previous_cn = className;
				}

				//end
			}
			_doc.Sections.RemoveAt(0);
			e.Result = _doc;
		}
		class MailMerge_MergeField : Aspose.Words.Reporting.IFieldMergingCallback
		{
			public void FieldMerging(Aspose.Words.Reporting.FieldMergingArgs args)
			{
				Aspose.Words.Reporting.FieldMergingArgs e = args;
				if (e.FieldName == "QRCODE")
				{
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
				_doc.Sections.Clear();

				#region 班級學生的電子報表
				//建立一個學生電子報表
				//傳入參數 : 報表名稱,學年度,學期,類型(學生/班級/教師/課程)
				paperForStudent = new SmartSchool.ePaper.ElectronicPaper("學生家長代碼邀請函 for App", School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

				//學生個人的文件
				Document each_page = new Document(studentTemplate);
				MemoryStream stream = new MemoryStream();
				each_page.Save(stream, SaveFormat.Doc);
				_doc.Sections.Add(_doc.ImportNode(each_page.Sections[0], true)); //合併至doc

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
			//System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
			//sd.Title = "另存新檔";
			//sd.FileName = "合併欄位.doc";
			//sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
			//if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			//{
			//	try
			//	{
			//		Document document = new Document(Properties.Resources.)
			//		.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
			//		System.Diagnostics.Process.Start(sd.FileName);
			//	}
			//	catch (Exception ex)
			//	{
			//		FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
			//		return;
			//	}
			//}
		}
	}
}
