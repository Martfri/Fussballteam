# Automatic generation of web applications from Excel documents

The aim is to create a web application based on the MVC architecture that can interact with an Excel document that is available as a micro service. The user of the web application is able to call up the micro service by entering the inputs in an input mask. When the service is called, it writes the parameters into predefined fields of the Excel file, calculate the formulas in the file and return predefined cell values as a result and display them in the UI of the web application. On the other hand, changes in the Excel wil be automatically reflected in the UI.

# Usage

- Microsoft SQL Server needs to be installed. Alternatively, SQL Server Express LocalDB can be used for testing purposes.
- To scaffold a new database migration the following command needs to be executed in the root folder of the project with the .NET Core CLI: `dotnet ef migrations add InitialCreatee`
- The Migration can then be applied with the following command `dotnet ef database update`
- Open the final release, download "Fussballteam.zip", unzip and start the exe file (Windows) to start the web application.
- Login with different users to work with their corresponding data. Each User is assigned an excel sheet. For newly registered users there needs to be a manual set up of a new sheet according to the template. 
- Register a new user. User name needs to match the sheet name to be authorized to work with the excel sheet. E.g: Liverpool@localhost.com 

# Files

* Models/Player.cs: represents the player's data stored in the Excel file.
* Services/ExcelService.cs: offers interfaces for manipulating and reading the the sheet.

 
# Sources

Microsoft SQL Server:
https://www.microsoft.com/en-us/sql-server/sql-server-downloads

SQl Server Express LocalDB:
https://learn.microsoft.com/en-us/sql/database-engine/configure-windows/sql-server-express-localdb?view=sql-server-ver16

Database Migrations Overview:
https://learn.microsoft.com/en-us/ef/core/managing-schemas/migrations/?tabs=vs





