using FISCA.Permission;
using FISCA.Presentation;
using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InvitationParents
{
    public class Program
    {
		[FISCA.MainMethod]
		public static void Main() 
		{
			//學生頁面選取學生資料
			FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
			item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
			item1["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Enable = false;
			item1["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Click += delegate
			{
				Form form = new Form(K12.Presentation.NLDPanels.Student.SelectedSource);
				form.ShowDialog();
			};

			//班級頁面選取班級資料
			FISCA.Presentation.RibbonBarItem item2 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
			item2["報表"].Image = Properties.Resources.Report;
			item2["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
			item2["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Enable = false;
			item2["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Click += delegate
			{
				List<StudentRecord> studs = K12.Data.Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
				List<string> strStuds = new List<string>();
				foreach (StudentRecord sr in studs) {
					if (strStuds.Contains(sr.ID)) {
						strStuds.Add(sr.ID);
					}
				}

				Form form = new Form(strStuds);
				form.ShowDialog();
			};

			//權限設定(學生)
			K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
			{
				if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.學生家長邀請函權限)
				{
					item1["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Enable = true;
				}
				else
					item1["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Enable = false;
			};

			////權限設定(班級)
			//K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
			//{
			//	if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.班級家長邀請函權限)
			//	{
			//		item1["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Enable = true;
			//	}
			//	else
			//		item1["報表"]["代碼相關報表"]["學生家長代碼邀請函 for App"].Enable = false;
			//};

			//權限設定
			Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
			permission.Add(new RibbonFeature(Permissions.學生家長邀請函, "學生家長代碼邀請函 for App"));
			//permission = RoleAclSource.Instance["班級"]["功能按鈕"];
			//permission.Add(new RibbonFeature(Permissions.班級家長邀請函, "學生家長代碼邀請函 for App"));
		}
    }
}
