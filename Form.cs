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
		private List<string> studentIds;
		private List<string> classIds;
		private QueryHelper queryHelper;

		public Form(List<string> selectStudentIds)
		{
			InitializeComponent();
			studentIds = selectStudentIds;
			this.queryHelper = new QueryHelper();
		}

		private void buttonX1_Click(object sender, EventArgs e)
		{
			BackgroundWorker bgw = new BackgroundWorker();
			bgw = new BackgroundWorker();
			bgw.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
			bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
			bgw.RunWorkerAsync(studentIds);
			//bgw.RunWorkerAsync();
		}
		private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			List<string> student = (List<string>)e.Argument;

			Dictionary<string, StudentRecord> dsr = K12.Data.Student.SelectByIDs(student).ToDictionary(x => x.ID, x => x);
			string DSNSName = FISCA.Authentication.DSAServices.AccessPoint;

			string schoolName = K12.Data.School.ChineseName;
			string ids = string.Join(",", dsr.Keys);

			//DataTable dt = queryHelper.Select("select student.id,student.parent_code from student where student.parent_code <> '' ");
			//DataTable dt = queryHelper.Select("select student.id,student.parent_code from student join class.id = student.ref_class_id  order by class.display_order,class.name,student.seat_no");
			DataTable dt = queryHelper.Select(@"select student.id,student.parent_code from student
join class on class.id = student.ref_class_id where student.id in (" + ids + ")order by class.grade_year,class.display_order,class.class_name,student.seat_no");
			Document template = new Document(new MemoryStream(Properties.Resources.QRcode));
			Document doc = new Document();
			Dictionary<string, object> merge = new Dictionary<string, object>();
			foreach (DataRow row in dt.Rows)
			{
				Document perPage = template.Clone();
				string parentsCode = "";
				string studentId = "";
				merge.Clear();
				studentId = row["id"] + "";
				parentsCode = row["parent_code"] + "";

				StudentRecord sr;
				if (dsr.ContainsKey(studentId))
				{
					sr = dsr[studentId];

					if (sr.Class != null)
					{
						merge.Add("年級", sr.Class.GradeYear);
						merge.Add("班級名稱", sr.Class.Name);
					}
					merge.Add("學校名稱", schoolName);
					merge.Add("座號", sr.SeatNo);
					merge.Add("學生姓名", sr.Name);
					merge.Add("QRCODE", parentsCode + "@" + DSNSName);
					merge.Add("家長代碼", parentsCode);
				}

				perPage.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
				perPage.MailMerge.Execute(merge.Keys.ToArray<string>(), merge.Values.ToArray<object>());
				perPage.MailMerge.RemoveEmptyParagraphs = true;
				perPage.MailMerge.DeleteFields();

				doc.Sections.Add(doc.ImportNode(perPage.Sections[0], true));
			}

			doc.Sections.RemoveAt(0);
			e.Result = doc;
		}
		void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
		{
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

		void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{

			if (e.Error != null)
			{
				FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤:" + e.Error.Message);
				return;
			}

			Document document = (Document)e.Result;
			string inputReportName = "QRCODE";
			string reportName = inputReportName;

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


		/// <summary>
		/// 學生電子報表
		/// </summary>
		SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }

		private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
		{
			if (checkBoxX1.ThreeState == true) 
			{
				#region 班級學生的電子報表
				//建立一個學生電子報表
				//傳入參數 : 報表名稱,學年度,學期,類型(學生/班級/教師/課程)
				paperForStudent = new SmartSchool.ePaper.ElectronicPaper("學生家長代碼邀請函 for App", School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

				//學生個人的文件
				Document each_page = new Document(, "", LoadFormat.Doc, "");
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
	}
}
