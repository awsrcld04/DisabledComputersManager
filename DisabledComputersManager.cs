using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using Microsoft.Win32;
using System.Management;
using System.Reflection;
using System.Diagnostics;

namespace DisabledComputersManager
{
    class DCMMain
    {
        struct DisabledComputersParams
        {
            public string strDisabledComputersLocation;
            //public string strSendActionList;
            //public string strActionListSender;
            //public string strActionListRecipient;
            public int intDisabledHoldPeriod;
            public List<string> lstExcludePrefix;
            public List<string> lstExcludeSuffix;
            public List<string> lstExclude;
        }

        struct CMDArguments
        {
            public bool bParseCmdArguments;
        }

        static bool funcLicenseCheck()
        {
            try
            {
                string strLicenseString = "";
                bool bValidLicense = false;

                TextReader tr = new StreamReader("sotfwlic.dat");

                try
                {
                    strLicenseString = tr.ReadLine();

                    if (strLicenseString.Length > 0 & strLicenseString.Length < 29)
                    {
                        // [DebugLine] Console.WriteLine("if: " + strLicenseString);
                        Console.WriteLine("Invalid license");

                        tr.Close(); // close license file

                        return bValidLicense;
                    }
                    else
                    {
                        tr.Close(); // close license file
                        // [DebugLine] Console.WriteLine("else: " + strLicenseString);

                        string strMonthTemp = ""; // to convert the month into the proper number
                        string strDate;

                        //Month
                        strMonthTemp = strLicenseString.Substring(7, 1);
                        if (strMonthTemp == "A")
                        {
                            strMonthTemp = "10";
                        }
                        if (strMonthTemp == "B")
                        {
                            strMonthTemp = "11";
                        }
                        if (strMonthTemp == "C")
                        {
                            strMonthTemp = "12";
                        }
                        strDate = strMonthTemp;

                        //Day
                        strDate = strDate + "/" + strLicenseString.Substring(16, 1);
                        strDate = strDate + strLicenseString.Substring(6, 1);

                        // Year
                        strDate = strDate + "/" + strLicenseString.Substring(24, 1);
                        strDate = strDate + strLicenseString.Substring(4, 1);
                        strDate = strDate + strLicenseString.Substring(1, 2);

                        // [DebugLine] Console.WriteLine(strDate);
                        // [DebugLine] Console.WriteLine(DateTime.Today.ToString());
                        DateTime dtLicenseDate = DateTime.Parse(strDate);
                        // [DebugLine]Console.WriteLine(dtLicenseDate.ToString());

                        if (dtLicenseDate >= DateTime.Today)
                        {
                            bValidLicense = true;
                        }
                        else
                        {
                            Console.WriteLine("License expired.");
                        }

                        return bValidLicense;
                    }

                } //end of try block on tr.ReadLine

                catch
                {
                    // [DebugLine] Console.WriteLine("catch on tr.Readline");
                    Console.WriteLine("Invalid license");
                    tr.Close();
                    return bValidLicense;

                } //end of catch block on tr.ReadLine

            } // end of try block on new StreamReader("sotfwlic.dat")

            catch (System.Exception ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());

                // [DebugLine] System.Console.WriteLine(ex.Message);

                if (ex.Message.StartsWith("Could not find file"))
                {
                    Console.WriteLine("License file not found.");
                }
                else
                {
                    MethodBase mb1 = MethodBase.GetCurrentMethod();
                    funcGetFuncCatchCode(mb1.Name, ex);
                }

                return false;

            } // end of catch block on new StreamReader("sotfwlic.dat")
        }

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("A parameter is missing or is incorrect.");
            Console.WriteLine("Run DisabledComputersManager -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("DisabledComputersManager v1.0 (c) 2011 SystemsAdminPro.com");
            Console.WriteLine();
            Console.WriteLine("Description: Manage disabled computers");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following required parameters in the following order:");
            Console.WriteLine("-run                     required parameter");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("DisabledComputersManager -run");
        }

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                if (cmdargs[0] == "-run" & cmdargs.Length == 1)
                {
                    objCMDArguments.bParseCmdArguments = true;
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static DisabledComputersParams funcParseConfigFile(CMDArguments objCMDArguments2)
        {
            DisabledComputersParams newParams = new DisabledComputersParams();

            try
            {
                newParams.lstExclude = new List<string>();
                newParams.lstExcludePrefix = new List<string>();
                newParams.lstExcludeSuffix = new List<string>();

                newParams.intDisabledHoldPeriod = 0; //initialize

                // find all domain controllers & automatically exclude all domain controllers
                Domain dmThis = Domain.GetCurrentDomain();
                DomainControllerCollection dcCollection = dmThis.FindAllDomainControllers();
                foreach (DomainController dcThis in dcCollection)
                {
                    DirectoryEntry dcDE = dcThis.GetDirectoryEntry();
                    //[DebugLine] Console.WriteLine(dcDE.Properties["name"].Value.ToString());
                    if (dcDE != null)
                        newParams.lstExclude.Add(dcDE.Properties["name"].Value.ToString());
                }

                TextReader trConfigFile = new StreamReader("configDisabledComputersManager.txt");

                using (trConfigFile)
                {
                    string strNewLine = "";

                    while ((strNewLine = trConfigFile.ReadLine()) != null)
                    {
                        if (strNewLine.StartsWith("DisabledComputersLocation=") & strNewLine != "DisabledComputersLocation=")
                        {
                            newParams.strDisabledComputersLocation = strNewLine.Substring(26);
                            //[DebugLine] Console.WriteLine(newParams.strDisabledComputersLocation);
                        }
                        //if (strNewLine.StartsWith("SendActionList="))
                        //{
                        //    newParams.strSendActionList = strNewLine.Substring(15);
                        //    //[DebugLine] Console.WriteLine(newParams.strSendActionList);
                        //}
                        //if (strNewLine.StartsWith("ActionListSender="))
                        //{
                        //    newParams.strActionListSender = strNewLine.Substring(17);
                        //    //[DebugLine] Console.WriteLine(newParams.strActionListSender);
                        //}
                        //if (strNewLine.StartsWith("ActionListRecipient="))
                        //{
                        //    newParams.strActionListRecipient = strNewLine.Substring(20);
                        //    //[DebugLine] Console.WriteLine(newParams.strActionListRecipient);
                        //}
                        if (strNewLine.StartsWith("ExcludePrefix=") & strNewLine != "ExcludePrefix=")
                        {
                            newParams.lstExcludePrefix.Add(strNewLine.Substring(14));
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(14));
                        }
                        if (strNewLine.StartsWith("ExcludeSuffix=") & strNewLine != "ExcludeSuffix=")
                        {
                            newParams.lstExcludeSuffix.Add(strNewLine.Substring(14));
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(14));
                        }
                        if (strNewLine.StartsWith("Exclude=") & strNewLine != "Exclude=")
                        {
                            newParams.lstExclude.Add(strNewLine.Substring(8));
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(8));
                        }
                        if (strNewLine.StartsWith("DisabledHoldPeriod=") & strNewLine != "DisabledHoldPeriod=")
                        {
                            newParams.intDisabledHoldPeriod = Int32.Parse(strNewLine.Substring(19));
                            //[DebugLine] Console.WriteLine(strNewLine.Substring(19) + newParams.intDisabledHoldPeriod.ToString());
                        }
                    }
                }

                //[DebugLine] Console.WriteLine("# of Exclude= : {0}", newParams.lstExclude.Count.ToString());
                //[DebugLine] Console.WriteLine("# of ExcludePrefix= : {0}", newParams.lstExcludePrefix.Count.ToString());

                trConfigFile.Close();

                if (newParams.intDisabledHoldPeriod == 0)
                {
                    newParams.intDisabledHoldPeriod = 7;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            return newParams;
        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                Domain currentDomain = Domain.GetComputerDomain();
                if (currentDomain.DomainMode != DomainMode.Windows2000MixedDomain &
                    currentDomain.DomainMode != DomainMode.Windows2000NativeDomain &
                    currentDomain.DomainMode != DomainMode.Windows2003InterimDomain)
                {

                    // [DebugLine] Console.WriteLine("Entering funcProgramExecution");
                    if (funcCheckForFile("configDisabledAccountsManager.txt"))
                    {
                        DisabledComputersParams newParams = funcParseConfigFile(objCMDArguments2);
                        
                        funcToEventLog("DisabledComputersManager", "DisabledComputersManager started.", 100);

                        funcProgramRegistryTag("DisabledComputersManager");

                        funcModifyDisabledComputers(newParams);

                        //funcMoveDisabledComputers(newParams);

                        funcRemoveComputers(newParams);

                        funcToEventLog("DisabledComputersManager", "DisabledComputersManager stopped.", 101);
                    }
                    else
                    {
                        Console.WriteLine("Config file configDisabledComputersManager.txt could not be found.");
                    }
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static void funcProgramRegistryTag(string strProgramName)
        {
            try
            {
                string strRegistryProfilesPath = "SOFTWARE";
                RegistryKey objRootKey = Microsoft.Win32.Registry.LocalMachine;
                RegistryKey objSoftwareKey = objRootKey.OpenSubKey(strRegistryProfilesPath, true);
                RegistryKey objSystemsAdminProKey = objSoftwareKey.OpenSubKey("SystemsAdminPro", true);
                if (objSystemsAdminProKey == null)
                {
                    objSystemsAdminProKey = objSoftwareKey.CreateSubKey("SystemsAdminPro");
                }
                if (objSystemsAdminProKey != null)
                {
                    if (objSystemsAdminProKey.GetValue(strProgramName) == null)
                        objSystemsAdminProKey.SetValue(strProgramName, "1", RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static DirectorySearcher funcCreateDSSearcher()
        {
            try
            {
                System.DirectoryServices.DirectorySearcher objDSSearcher = new DirectorySearcher();
                // [Comment] Get local domain context

                string rootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                rootDSE = objrootDSESearcher.SearchRoot.Path;
                //Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
                objDSSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                //Console.WriteLine(objDSSearcher.SearchRoot.Path);

                return objDSSearcher;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }
        }

        static PrincipalContext funcCreatePrincipalContext()
        {
            PrincipalContext newctx = new PrincipalContext(ContextType.Machine);

            try
            {
                //Console.WriteLine("Entering funcCreatePrincipalContext");
                Domain objDomain = Domain.GetComputerDomain();
                string strDomain = objDomain.Name;
                DirectorySearcher tempDS = funcCreateDSSearcher();
                string strDomainRoot = tempDS.SearchRoot.Path.Substring(7);
                // [DebugLine] Console.WriteLine(strDomainRoot);
                // [DebugLine] Console.WriteLine(strDomainRoot);

                newctx = new PrincipalContext(ContextType.Domain,
                                    strDomain,
                                    strDomainRoot);

                // [DebugLine] Console.WriteLine(newctx.ConnectedServer);
                // [DebugLine] Console.WriteLine(newctx.Container);



                //if (strContextType == "Domain")
                //{

                //    PrincipalContext newctx = new PrincipalContext(ContextType.Domain,
                //                                    strDomain,
                //                                    strDomainRoot);
                //    return newctx;
                //}
                //else
                //{
                //    PrincipalContext newctx = new PrincipalContext(ContextType.Machine);
                //    return newctx;
                //}
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            if (newctx.ContextType == ContextType.Machine)
            {
                Exception newex = new Exception("The Active Directory context did not initialize properly.");
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, newex);
            }

            return newctx;
        }

        static bool funcCheckNameExclusion(string strName, DisabledComputersParams listParams)
        {
            try
            {
                bool bNameExclusionCheck = false;

                //List<string> listExclude = new List<string>();
                //listExclude.Add("Guest");
                //listExclude.Add("SUPPORT_388945a0");
                //listExclude.Add("krbtgt");

                if (listParams.lstExclude.Contains(strName))
                    bNameExclusionCheck = true;

                //string strMatch = listExclude.Find(strName);
                foreach (string strNameTemp in listParams.lstExcludePrefix)
                {
                    if (strName.StartsWith(strNameTemp))
                    {
                        bNameExclusionCheck = true;
                        break;
                    }
                }

                foreach (string strNameTemp in listParams.lstExcludeSuffix)
                {
                    if (strName.EndsWith(strNameTemp))
                    {
                        bNameExclusionCheck = true;
                        break;
                    }
                }

                return bNameExclusionCheck;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcMoveDisabledComputers(DisabledComputersParams currentParams)
        {
            try
            {
                PrincipalContext currentctx = funcCreatePrincipalContext();

                ComputerPrincipal newComputerPrincipal = new ComputerPrincipal(currentctx);

                PrincipalSearcher ps = new PrincipalSearcher(newComputerPrincipal);

                // Create an in-memory user object to use as the query example.
                ComputerPrincipal c = new ComputerPrincipal(currentctx);

                c.Enabled = false;

                // Tell the PrincipalSearcher what to search for.
                ps.QueryFilter = c;

                // Run the query. The query locates users 
                // that match the supplied user principal object. 
                PrincipalSearchResult<Principal> results = ps.FindAll();

                int intSearchResults = results.Count<Principal>();

                if (intSearchResults > 0)
                {

                    if (funcCheckForOU(currentParams.strDisabledComputersLocation))
                    {
                        TextWriter twCurrent = funcOpenOutputLog();
                        string strOutputMsg = "";

                        DirectoryEntry orgDE = new DirectoryEntry("LDAP://" + currentParams.strDisabledComputersLocation);

                        //[DebugLine] Console.WriteLine(orgDE.Path);
                        funcWriteToOutputLog(twCurrent, orgDE.Path);
                        //[DebugLine] Console.WriteLine();
                        //[DebugLine] funcWriteToOutputLog(twCurrent, "");

                        foreach (ComputerPrincipal cp in results)
                        {
                            if (!funcCheckNameExclusion(cp.Name, currentParams))
                            {
                                //[DebugLine] Console.WriteLine("Name: {0}", up.Name);
                                strOutputMsg = "Name: " + cp.Name;
                                funcWriteToOutputLog(twCurrent, strOutputMsg);

                                DirectoryEntry newDE = new DirectoryEntry("LDAP://" + cp.DistinguishedName);
                                //[DebugLine] Console.WriteLine("Path: {0}", newDE.Path);

                                strOutputMsg = "Path: " + newDE.Path;
                                funcWriteToOutputLog(twCurrent, strOutputMsg);

                                if (!newDE.Path.Contains(currentParams.strDisabledComputersLocation))
                                {
                                    //Console.WriteLine("Move to DisabledObjects");
                                    funcWriteToOutputLog(twCurrent, "Move to DisabledObjects");
                                    newDE.MoveTo(orgDE);
                                    newDE.CommitChanges();
                                    //[DebugLine] Console.WriteLine("Path: {0}", newDE.Path);
                                    strOutputMsg = "Path: " + newDE.Path;
                                    funcWriteToOutputLog(twCurrent, strOutputMsg);
                                    newDE.Close();
                                    //[DebugLine] Console.WriteLine();
                                    //[DebugLine] funcWriteToOutputLog(twCurrent, "");
                                }
                            }
                        }

                        funcCloseOutputLog(twCurrent);
                    }
                    else
                    {
                        Exception newex = new Exception("The OU path for disabled accounts is invalid.");
                        throw newex;
                    }
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcModifyDisabledComputers(DisabledComputersParams currentParams)
        {
            try
            {
                if (funcCheckForOU(currentParams.strDisabledComputersLocation))
                {
                    Domain thisDomain = Domain.GetComputerDomain();
                    string strDomainName = thisDomain.Name;
                    PrincipalContext ctxDisabledComputers = new PrincipalContext(ContextType.Domain, strDomainName,
                                                                             currentParams.strDisabledComputersLocation);

                    // Create the principal context for the usr object.
                    PrincipalContext ctx = funcCreatePrincipalContext();

                    // Create a PrincipalSearcher object.
                    PrincipalSearcher ps = new PrincipalSearcher();

                    // Create the principal user object from the context
                    ComputerPrincipal computer = new ComputerPrincipal(ctx);

                    computer.Enabled = false;

                    ps.QueryFilter = computer;

                    PrincipalSearchResult<Principal> psr = ps.FindAll();

                    funcToEventLog("DisabledComputersManager", "Number of disabled computers to process: " +
                                    psr.Count<Principal>().ToString(), 1001);

                    if (psr.Count<Principal>() > 0)
                    {
                        TextWriter twCurrent = funcOpenOutputLog();
                        string strOutputMsg = "";

                        funcWriteToOutputLog(twCurrent, "--------DisabledComputersManager: Processing disabled computers");

                        List<string> lstDisabledComputers = new List<string>();

                        foreach (ComputerPrincipal c in psr)
                        {
                            if (!funcCheckNameExclusion(c.Name, currentParams))
                            {
                                if (!c.DistinguishedName.Contains(currentParams.strDisabledComputersLocation))
                                {
                                    lstDisabledComputers.Add(c.Sid.ToString());

                                    //[DebugLine] Console.WriteLine("Account to Disable: {0}", c.Name);
                                    strOutputMsg = "Disabled computer: " + c.Name;
                                    funcWriteToOutputLog(twCurrent, strOutputMsg);

                                    //[DebugLine] Console.WriteLine(c.DistinguishedName);
                                    funcWriteToOutputLog(twCurrent, c.DistinguishedName);

                                    List<string> grplist = new List<string>();
                                    foreach (GroupPrincipal gp in c.GetGroups())
                                    {
                                        //[DebugLine] Console.WriteLine("{0} \t {1}", gp.Name, gp.Sid.ToString());
                                        grplist.Add(gp.Sid.ToString());
                                        //[DebugLine] funcWriteToOutputLog(twCurrent, gp.Name + "\t" + gp.Sid.ToString());
                                    }

                                    foreach (string strGrpSID in grplist)
                                    {
                                        funcRemoveComputerFromGroup(ctx, strGrpSID, c.Sid.ToString(), twCurrent);
                                    }

                                    c.Save(ctxDisabledComputers);
                                    ComputerPrincipal cp = ComputerPrincipal.FindByIdentity(ctxDisabledComputers, IdentityType.SamAccountName, c.SamAccountName);
                                    if (cp.DistinguishedName.Contains(currentParams.strDisabledComputersLocation))
                                    {
                                        strOutputMsg = "Computer: " + cp.Name + " - successfully moved to " + currentParams.strDisabledComputersLocation;
                                        funcWriteToOutputLog(twCurrent, strOutputMsg);
                                    }
                                }
                            }
                        }

                        if (lstDisabledComputers.Count > 0)
                        {
                            // Create group - ToBeDeleted[Date]
                            // Add accounts to the group for the day that this process runs
                            string strGroupName = "ComputersToBeDeleted" + DateTime.Today.ToLocalTime().ToString("MMddyyyy");
                            GroupPrincipal newgrp = GroupPrincipal.FindByIdentity(ctxDisabledComputers, strGroupName);

                            if (newgrp == null)
                            {
                                newgrp = new GroupPrincipal(ctxDisabledComputers, strGroupName);
                                newgrp.Description = "Computers";
                                newgrp.Save();
                            }

                            string strGroupMembersMessage = lstDisabledComputers.Count.ToString() + " computers to be added to group " + newgrp.Name;
                            funcToEventLog("DisabledComputersManager", strGroupMembersMessage, 1002);

                            foreach (string computerSID in lstDisabledComputers)
                            {
                                ComputerPrincipal c2 = ComputerPrincipal.FindByIdentity(ctxDisabledComputers, IdentityType.Sid, computerSID);
                                newgrp.Members.Add(c2);
                                newgrp.Save();
                            }
                        }

                        funcCloseOutputLog(twCurrent);
                    }
                }
                else
                {
                    Exception newex = new Exception("The OU path for disabled computers is invalid.");
                    throw newex;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static void funcRemoveComputerFromGroup(PrincipalContext currentctx, string grpSID, string computerSID, TextWriter twCurrent)
        {
            try
            {
                ComputerPrincipal computer = ComputerPrincipal.FindByIdentity(currentctx, IdentityType.Sid, computerSID);
                DirectoryEntry tempDE = new DirectoryEntry("LDAP://" + computer.DistinguishedName);
                string strPrimaryGroupID = tempDE.Properties["primaryGroupID"].Value.ToString();
                if (!grpSID.EndsWith("-" + strPrimaryGroupID))
                {
                    GroupPrincipal grp = GroupPrincipal.FindByIdentity(currentctx, IdentityType.Sid, grpSID);
                    //[DebugLine] Console.WriteLine("Remove computer from group {0}", grp.Name);
                    funcWriteToOutputLog(twCurrent, "Remove computer from group " + grp.Name);
                    if (computer != null & grp != null)
                    {
                        grp.Members.Remove(computer);
                    }
                    grp.Save();
                }
                else
                {
                    //[DebugLine] Console.WriteLine("Do not remove computer from primary group.");
                    funcWriteToOutputLog(twCurrent, "Do not remove computer from primary group.");
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcRemoveComputers(DisabledComputersParams currentParams)
        {
            try
            {
                DateTime dtHoldPeriod = DateTime.Today.AddDays(-currentParams.intDisabledHoldPeriod);

                Domain thisDomain = Domain.GetComputerDomain();
                string strDomainName = thisDomain.Name;
                PrincipalContext ctxDisabledComputers = new PrincipalContext(ContextType.Domain, strDomainName,
                                                                         currentParams.strDisabledComputersLocation);

                TextWriter twCurrent = funcOpenOutputLog();

                // Create a PrincipalSearcher object.
                PrincipalSearcher ps = new PrincipalSearcher();

                // Create an in-memory user object to use as the query example.
                GroupPrincipal grp = new GroupPrincipal(ctxDisabledComputers);

                grp.Description = "Computers";

                // Tell the PrincipalSearcher what to search for.
                ps.QueryFilter = grp;

                PrincipalSearchResult<Principal> psr = ps.FindAll();

                string strEventLogMsg = "";
                string strOutputMsg = "";

                if (psr.Count<Principal>() > 0)
                {
                    strEventLogMsg = "Number of ComputersToBeDeleted groups found: " + psr.Count<Principal>().ToString();
                    funcToEventLog("DisabledComputersManager", strEventLogMsg, 1003);

                    foreach (GroupPrincipal g in psr)
                    {
                        string strDateString = g.Name.Substring(20, 2).ToString() + "/" +
                                               g.Name.Substring(22, 2).ToString() + "/" +
                                               g.Name.Substring(24).ToString();

                        //[DebugLine] Console.WriteLine(strDateString);

                        DateTime dtGroupCreated = Convert.ToDateTime(strDateString);
                        if (dtGroupCreated < dtHoldPeriod)
                        {
                            strEventLogMsg = "Number of computers to be deleted in group " + g.Name + ": " + g.Members.Count.ToString();
                            funcToEventLog("DisabledComputersManager", strEventLogMsg, 1004);

                            foreach (UserPrincipal p in g.GetMembers())
                            {
                                strOutputMsg = "Deleting computer: " + p.Name;
                                funcWriteToOutputLog(twCurrent, strOutputMsg);
                                p.Delete();
                            }

                            if (g.Members.Count == 0)
                            {
                                strOutputMsg = "Deleting group: " + g.Name;
                                funcWriteToOutputLog(twCurrent, strOutputMsg);
                                g.Delete();
                            }
                        }
                        else
                        {
                            //group is within hold period
                            strEventLogMsg = "No computers to be deleted from group " + g.Name + " - group is within hold period ";
                            funcToEventLog("DisabledComputersManager", strEventLogMsg, 1004);
                        }
                    }
                }
                else
                {
                    strEventLogMsg = "No ComputersToBeDeleted groups were found";
                    funcToEventLog("DisabledComputersManager", strEventLogMsg, 1003);
                }

                funcCloseOutputLog(twCurrent);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            try
            {
                string strLogName;

                strLogName = "Application";

                if (!EventLog.SourceExists(strAppName))
                    EventLog.CreateEventSource(strAppName, strLogName);

                //EventLog.WriteEntry(strAppName, strEventMsg);
                EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcCheckForFile", "f0");
            dCatchTable.Add("funcCheckForOU", "f1");
            dCatchTable.Add("funcCheckNameExclusion", "f2");
            dCatchTable.Add("funcCloseOutputLog", "f3");
            dCatchTable.Add("funcCreateDSSearcher", "f4");
            dCatchTable.Add("funcCreatePrincipalContext", "f5");
            dCatchTable.Add("funcErrorToEventLog", "f6");
            dCatchTable.Add("funcGetFuncCatchCode", "f7");
            dCatchTable.Add("funcLicenseActivation", "f8");
            dCatchTable.Add("funcLicenseCheck", "f9");
            dCatchTable.Add("funcModifyDisabledComputers", "f10");
            dCatchTable.Add("funcMoveDisabledComputers", "f11");
            dCatchTable.Add("funcOpenOutputLog", "f12");
            dCatchTable.Add("funcParseCmdArguments", "f13");
            dCatchTable.Add("funcParseConfigFile", "f14");
            dCatchTable.Add("funcPrintParameterSyntax", "f15");
            dCatchTable.Add("funcPrintParameterWarning", "f16");
            dCatchTable.Add("funcProgramExecution", "f17");
            dCatchTable.Add("funcProgramRegistryTag", "f18");
            dCatchTable.Add("funcRemoveComputerFromGroup", "f19");
            dCatchTable.Add("funcRemoveComputers", "f20");
            dCatchTable.Add("funcToEventLog", "f21");
            dCatchTable.Add("funcWriteToErrorLog", "f22");
            dCatchTable.Add("funcWriteToOutputLog", "f23");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);
            funcErrorToEventLog("DisabledComputersManager");
        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                string strPath = Directory.GetCurrentDirectory();

                if (!Directory.Exists(strPath + "\\Log"))
                {
                    Directory.CreateDirectory(strPath + "\\Log");
                    if (Directory.Exists(strPath + "\\Log"))
                    {
                        strPath = strPath + "\\Log";
                    }
                }
                else
                {
                    strPath = strPath + "\\Log";
                }

                FileStream newFileStream = new FileStream(strPath + "\\Err-DisabledComputersManager.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static TextWriter funcOpenOutputLog()
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat2 = "MMddyyyy"; // for log file directory creation

                string strPath = Directory.GetCurrentDirectory();

                if (!Directory.Exists(strPath + "\\Log"))
                {
                    Directory.CreateDirectory(strPath + "\\Log");
                    if (Directory.Exists(strPath + "\\Log"))
                    {
                        strPath = strPath + "\\Log";
                    }
                }
                else
                {
                    strPath = strPath + "\\Log";
                }

                string strLogFileName = strPath + "\\DisabledComputersManager" + dtNow.ToString(dtFormat2) + ".log";

                FileStream newFileStream = new FileStream(strLogFileName, FileMode.Append, FileAccess.Write);
                TextWriter twOuputLog = new StreamWriter(newFileStream);

                return twOuputLog;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }

        }

        static void funcWriteToOutputLog(TextWriter twCurrent, string strOutputMessage)
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                //string dtFormat = "MM/dd/yyyy";
                string dtFormat2 = "MMddyyyy HH:mm";
                //string dtFormat3 = "MMddyyyy HH:mm:ss";

                twCurrent.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat2), strOutputMessage);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCloseOutputLog(TextWriter twCurrent)
        {
            try
            {
                twCurrent.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcErrorToEventLog(string strAppName)
        {
            string strLogName;

            strLogName = "Application";

            if (!EventLog.SourceExists(strAppName))
                EventLog.CreateEventSource(strAppName, strLogName);

            //EventLog.WriteEntry(strAppName, strEventMsg);
            EventLog.WriteEntry(strAppName, "An error has occured. Check log file.", EventLogEntryType.Error, 0);
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    funcPrintParameterWarning();
                }
                else
                {
                    if (args[0] == "-?")
                    {
                        funcPrintParameterSyntax();
                    }
                    else
                    {
                        string[] arrArgs = args;
                        CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                        if (objArgumentsProcessed.bParseCmdArguments)
                        {
                            funcProgramExecution(objArgumentsProcessed);
                        }
                        else
                        {
                            funcPrintParameterWarning();
                        } // check objArgumentsProcessed.bParseCmdArguments
                    } // check args[0] = "-?"
                } // check args.Length == 0
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }
    }
}
