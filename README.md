## SpringBoot project for Excel processing


### This SpringBoot project performs two actions on Excel files:

### Excel to JSON conversion:
- This functionality reads data from an Excel sheet and converts it into a JSON object or array.
- The specific logic and output format will depend on your implementation.
  
### Excel sheet splitting: 
- This functionality reads data from an Excel sheet and separates it based on a defined condition (e.g., column value).
- The resulting data is then written to two new Excel sheets.

## Getting started
### Prerequisites:
-Java 11 or greater than this version
-Maven

## Project Structure:

- src/main/java: Contains Java code for the application.
- src/main/resources: Contains configuration files and resources.
- pom.xml: Defines dependencies and build configuration.
- README.md: This file (you're reading it right now!)

### Clone the repository:

 - git clone <your_repository_url>
 - Build the project:
 - mvn clean install
 - Run the application:c

##### The application accepts arguments through Spring Boot profiles. You can activate specific profiles based on your desired functionality:

excel-to-json: Converts the specified Excel sheet to JSON.
excel-split: Splits the specified Excel sheet into two based on your criteria.


For example, to convert the sheet "Sheet1" of the file "data.xlsx" to JSON, run:

java -jar target/your-project-name.jar --spring.profiles.active=excel-to-json path/to/data.xlsx Sheet1

To split the sheet "Sheet1" of the file "data.xlsx" based on a column named "Category", with output files named "category1.xlsx" and "category2.xlsx", run:

java -jar target/your-project-name.jar --spring.profiles.active=excel-split path/to/data.xlsx Sheet1 Category
