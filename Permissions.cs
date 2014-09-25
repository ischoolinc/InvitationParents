using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InvitationParents
{
	class Permissions
	{
		public static bool 學生家長邀請函權限
		{
			get
			{
				return FISCA.Permission.UserAcl.Current[學生家長邀請函].Executable;
			}
		}

		public static bool 班級家長邀請函權限
		{
			get
			{
				return FISCA.Permission.UserAcl.Current[班級家長邀請函].Executable;
			}
		}

		public static string 學生家長邀請函 = "InvitationParents-{E3E7C850-2B23-454D-80EE-621469205118}";
		public static string 班級家長邀請函 = "InvitationParents-{796F477A-BF52-4054-94D1-0BA4E1FA8585}";
	}
}
