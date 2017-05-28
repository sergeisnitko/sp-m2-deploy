using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class Security
    {
        public static void RemoveGroupPermissions(this List list, Group group)
        {
            ClientContext clientContext = (ClientContext)list.Context;

            if (!list.IsPropertyAvailable("HasUniqueRoleAssignments"))
                clientContext.Load(list, currentList => currentList.HasUniqueRoleAssignments);

            if (!group.IsPropertyAvailable("Id"))
                clientContext.Load(group, currentGroup => currentGroup.Id);

            GroupCollection listGroups = list.RoleAssignments.Groups;
            clientContext.Load(listGroups,
                currentGroups => currentGroups.Include(
                    currentGroup => currentGroup.Id)
            );
            clientContext.ExecuteQuery();

            Group groupInList = listGroups.Cast<Group>().FirstOrDefault(currentGroup => currentGroup.Id == group.Id);

            if (groupInList != null)
            {
                if (!list.HasUniqueRoleAssignments)
                {
                    list.BreakRoleInheritance(true, false);
                    list.Update();
                }

                listGroups.RemoveById(group.Id);
                clientContext.ExecuteQuery();
            }
        }

        public static string RemoveSecurityTokens(this string LoginName)
        {
            return LoginName.Replace("i:0#.w|", "").Trim();
        }
        public static string AddSecurityTokens(this string LoginName)
        {
            return ("i:0#.w|" + LoginName).Trim();
        }
    }
}
