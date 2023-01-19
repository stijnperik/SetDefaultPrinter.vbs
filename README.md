# SetDefaultPrinter.vbs

------------------------------------------------------------------------------------
VBScript to set a default printer based on AD group membership
By stijnperik (2023)

USAGE:
- Set this as a user login script by group policy.
- Add a domain string (NETBIOS DOMAINNAME).
- Create a new AD group for each printer you want to be default.
- Under section 'Sub routine for creating Printer connections', add your printer share name matching your default printer group.
- Add users to the default printer groups.
- After login, the user will get their default printer based on their AD group membership.
		
VER: 0.5
-----------------------------------------------------------------------------------
