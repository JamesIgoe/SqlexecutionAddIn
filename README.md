# Sql Execution Add-In

A library that provides simplified threaded and non-threaded execution of SQL statements. It can be used to expose an interface to Microsoft Excel clients, providing threaded, async operations to a single-thread application. You will need to build this on a computer with the MS Office DLL's and/or PIA's. 


NuGet Package

* The NuGet package is avilable here: https://www.nuget.org/packages/SQLExecution/


SQL Execution Documentation (NuGet Package)

* 1.3.0: Added non-trusted execution parameters, corrected threaded execution
* 1.2.0: Made class virtual to allow instantiation
* 1.1.1: Corrected assembly information
* 1.1.0: Corrected class hierarchy

SQL Execution Tests

* See support the site, http://comparative-advantage.com/code/SQL_ExecutionHelp.php, for full usage syntax

VBA Usage

An example using the code to execute SQL asynchronously and write it out to different sheets. In effect, the execution time is nearer to the execution time of the slowest command object, rather than being the sum of execution times. 

Some Notes:

* The data classes used can be used to execute SQL, and either store data, or write to sheets
* Some fields in the code use global parameters as string for server instance, database name and timeout, prefaced by gstr
* Each item will execute independently, and do not need to be pointing at the same server/database
* See support the site, http://comparative-advantage.com/code/SQL_ExecutionHelp.php, for usage syntax

