# DgrepLdapReader
# Read me 1/29/2023
  # This script will convert AADDS's LDAP 1644 DGrep output into Excel pivot tables for workload analysis, to use this script:
  #    1. Run DGrep with 
        # source
        # | sort by TIMESTAMP asc
        # | project-rename LDAPServer=RoleInstance, TimeGenerated=PreciseTimeStamp, StartingNode=Data1, Filter=Data2, VisitedEntries=Data3, ReturnedEntries=Data4, Client=Data5, SearchScope=Data6, AttributeSelection=Data7, ServerControls=Data8, UsedIndexes=Data9, PagesReferenced=Data10, PagesReadFromDisk=Data11, PagesPreReadFromDisk=Data12, CleanPagesModified=Data13, DirtyPagesModified=Data14, SearchTimeMS=Data15, AttributesPreventingOptimization=Data16, User=Data17
        #     | extend ClientIP=split(Client,":",0)
        #     | extend ClientPort=split(Client,":",1)
        # | project LDAPServer, TimeGenerated, ClientIP, ClientPort, StartingNode, Filter, SearchScope, AttributeSelection, ServerControls, VisitedEntries, ReturnedEntries, UsedIndexes, PagesReferenced, PagesReadFromDisk, PagesPreReadFromDisk, CleanPagesModified, DirtyPagesModified, SearchTimeMS, AttributesPreventingOptimization, User
  #    2. Output to CSV
  #    3. Put CSV in same directory as this script.
  #    4. Script will perform string replacement in TimeGenerated, ClientIP, ClientPort fields, then calls Excel to import resulting CSV, create pivot tables for common ldap search analysis scenarios. 
  # Note: Script requires 64bits Excel.
  #
  # DgrepLdapReader.ps1 
    #		Steps: 
    #   	1. Put downloaded Dgrep Log.*.csv to same directory as DgrepLdapReader.ps1
    #   	2. Run script

  # Script info:    https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/event1644reader-analyze-ldap-query-performance
    #   Latest:       https://github.com/mingchen-script/DgrepLdapReader.ps1
    # AD Schema:      https://docs.microsoft.com/en-us/windows/win32/adschema/active-directory-schema
    # AD Attributes:  https://docs.microsoft.com/en-us/windows/win32/adschema/attributes
#------Script variables block, modify to fit your needs ---------------------------------------------------------------------
  $g_ColorBar   = $True                 # Can set to $false to speed up excel import & reduce memory requirement. 
  $g_ColorScale = $True                 # Can set to $false to speed up excel import & reduce memory requirement. Color Scale requires '$g_ColorBar = $True' for color index. 
