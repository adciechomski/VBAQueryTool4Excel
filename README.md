# VBAQueryTool4Excel
Tool designed to make SQL code customization from Excel UI and get desired report.

To use modules in your project trigger it with
Sub Initialization()

As of today module is compatible with SQL Microsoft Servers using SQL Server authentication.

Module is consuming your queries written in txt or sql files and process.
In this version you can allow user to customize filter attributes such as "LIKE", "IN", "NOT IN", "=".

To allow user to define filter set in query file use "--\\--" at the end of line. 
Code will recognize it and provide user with UserForm to enter desired parameters or even create options 
in ComboBox for list ["IN", "NOT IN"]

```
AND c.Country in ('Poland','Belgium') --\\--
```
or
```
c.ContactName LIKE 'M%' --\\--
```
or
```
c.ContactName = 'Maria' --\\--
```

As a result user will receive data inserted into new workSheet according to set customized filters.

*Module requires you to use aliases for tables in your queries.
