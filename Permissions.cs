using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InvitationParents
{
	class Permissions
	{
		public static bool 家長邀請函權限
		{
			get
			{
				return FISCA.Permission.UserAcl.Current[家長邀請函].Executable;
			}
		}

		public static string 家長邀請函 = "InvitationParents-{E566229B-8361-4918-BED0-D674ECF44DCC}";

	}
}
