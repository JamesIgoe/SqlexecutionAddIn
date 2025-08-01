# Sql Execution Add-In

A library that provides simplified threaded and non-threaded execution of SQL statements. It can be used to expose an interface to Microsoft Excel clients, providing threaded, async operations to a single-thread application. You will need to build this on a computer with the MS Office DLL's and/or PIA's. 

## NuGet Package

* The NuGet package is avilable here: https://www.nuget.org/packages/SQLExecution/

### SQL Execution Documentation (NuGet Package)

* 1.3.0: Added non-trusted execution parameters, corrected threaded execution
* 1.2.0: Made class virtual to allow instantiation
* 1.1.1: Corrected assembly information
* 1.1.0: Corrected class hierarchy

### SQL Execution Tests

* For usage and syntax see [SQL_ExecutionHelp.md](./SQL_ExecutionHelp.md)

### VBA Usage

An example using the code to execute SQL asynchronously and write it out to different sheets. In effect, the execution time is nearer to the execution time of the slowest command object, rather than being the sum of execution times. 

Some Notes:

* The data classes used can be used to execute SQL, and either store data, or write to sheets
* Some fields in the code use global parameters as string for server instance, database name and timeout, prefaced by gstr
* Each item will execute independently, and do not need to be pointing at the same server/database
* See support the site, http://comparative-advantage.com/code/SQL_ExecutionHelp.php, for usage syntax

## Gemini Suggested Improvements:

This C# library allows for asynchronous SQL execution and integration with VBA in Excel. Based on the provided information, here are some potential improvements:

* Error Handling and Robustness: The documentation doesn't specify how the library handles different types of SQL errors (e.g., connection failures, query syntax errors, timeouts) when executing multiple commands in parallel. A robust library would include a clear strategy for logging errors, handling individual command failures without stopping other executions, and providing detailed error information back to the caller.

* Modernization and Dependencies: Since the library is a few years old, it's worth evaluating if its dependencies are up-to-date and if there are newer .NET features or libraries that could simplify the code, improve performance, or add new functionality. For example, using async/await patterns more extensively could make the code cleaner and more idiomatic for modern C#.

* Documentation and Examples: While the project has some documentation, a more comprehensive guide with a wider range of examples would be beneficial. Examples could include different database providers, various SQL query types (e.g., SELECT, INSERT, UPDATE), and how to handle and process the results in VBA. The project could also benefit from a "Getting Started" guide to help new users quickly set it up.

* Unit and Integration Testing: The repository doesn't show any testing frameworks or test files. Implementing a suite of unit tests would ensure that individual components of the library work as expected, and integration tests would verify that the entire process, including the interaction with a database, works correctly. This would make the library more reliable and easier to maintain.

* Configuration and Customization: The current version might have hard-coded configurations or assumptions. Providing more customizable options, such as connection string management, configurable timeout values, or the ability to specify different output formats for the results, would make the library more versatile.
