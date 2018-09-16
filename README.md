# Description

This is a companion program for loading new rounds of testing data into the FIND
Interactive guide for high-quality malaria RDT selection. This companion program takes
an Excel file in a predefined format and converts it into a set of SQL insert statements
that can be executed against the database in order to import the new data. The companion 
program has been written using the Groovy language.  

# Disclaimer

This program is not a part of the FIND Interactive guide for high-quality malaria 
RDT selection and is not required neither for its operation nor for data import/export. 
The program is provided purely for convenience. You are free to use or not to use it.

We will not provide any further help, documentation and support for the program whatsover. 
We cannot hold liable for any potential damage that can occur from usage of this program.
Use it wisely for your convenience and at your own risk. 

# Excel File Format

The program expects that the Excel file has exactly the same format as was used for 
the Round 8. The technical documentation for the Interactive Guide contains an SQL query
that can be used to generate an Excel file in a proper format. Using the program on a file
with a different format will naturally lead to wrong results, you will have to modify 
the companion program if the data format changes.   

# Prerequisites

You will need Git and Java 8 installed on your machine.

# Usage

```
git clone https://github.com/rtitov73/find-data-loader.git
cd find-data-loader/
gradlew run -Dinput=<input_file> -Doutput=<output_file> [-DtableName=<tableName>]
```

Example:
```
git clone https://github.com/rtitov73/find-data-loader.git
cd find-data-loader/
gradlew run -Dinput="f:\R8.xlsx" -Doutput="f:\output.sql" -DtableName=malaria_rdt_tests_rnd_1_8
```

This command will take the data from the Excel file called **R8.xlsx**, convert it into
a set of SQL insert statements and store those statements in the **output.sql** file.
The table name will be **malaria_rdt_tests_rnd_1_8**

When you run the command for the first time it will load Gradle and required dependencies, so be patient.