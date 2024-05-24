# Equipment Assignment Contract Generator

This repository contains the Rails service `EquipmentAssignContractGenerationService`, which is used for generating equipment assignment contracts dynamically based on templates and employee data. The service fills out a Word document with specific equipment and employee details, and attaches the generated document to a given contract record.

## Overview

The `EquipmentAssignContractGenerationService` is designed to automate the creation of equipment loan contracts for employees. It utilizes a DOCX template and fills in details like equipment name, model, code, cost, and the assigned employee's details.
The Service requires users to manually configure DOCX templates and files with placeholders that have the following format: {{developer}} 



## Requirements

- Ruby on Rails
- Ruby gem `docx` (for handling DOCX files)

Make sure to add the `docx` gem to your Gemfile:

### Set-up 

Template Document: Place your Word document template in the public/contract_templates directory.
Initialization: Initialize the service with a contract object, employee name, employee object, and an array of active equipment associated with the contract.


###FUnctionality
- Logging Equipment Details: If equipment details are available, they are logged to the console.
- Document Generation: The service opens the DOCX template, replaces placeholders with actual data, and then saves the generated document.
- Handling Multiple Equipment: The service can handle multiple pieces of equipment, adding each to the contract document in a designated table.
- Employee Details: Specific employee details like RTN are substituted into the document.
- Saving and Attachment: The generated document is temporarily saved, attached to the contract object, and then the temporary file is deleted.


###Docuentation 
- fill_table_with_equipment: Fills a table in the document with equipment details.
- fill_equipment_row: Fills a single row of the equipment table with data.
- substitute_equipment_text: Substitutes placeholders in the equipment data table.
- substitute_employee_details: Substitutes employee-specific placeholders in the document.
- save_and_attach_document: Saves the filled document and attaches it to the contract record.



###License
This markdown content is ready for integration into your project's documentation. Adjustments can be made as needed to better fit your project's context or guidelines.
