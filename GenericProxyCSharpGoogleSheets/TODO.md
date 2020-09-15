
# CRUD
* ✓ Define the 4 CRUD methods as stubs
* Are Update and Delete needed for this project ?
  * Not for now - maybe later
* Add functionality
* ✓ Create
* ✓ Read
* Update - https://stackoverflow.com/questions/37462887/update-a-cell-with-c-sharp-and-sheets-api-v4
* Delete

# REST
* Add REST endpoint

# GUI
* Move project into a web project with GUI
* Input field
  * Should take root sheet ID
  * Editable
  * Should save to flat file
* Drop down
  * Should load configuration from root sheet
  * Should present loaded sheet IDs and sheet names
  * When option is chosen should present user with how to contact endpoint 

# Refactoring
* Move code into functions
* Tidy up main execution flow
* Write credentials.json to other location better suited for project so it can run on webserver
  * Can obj/Debug be used ?
  * https://stackoverflow.com/questions/57863574/google-sheets-api-v4-not-working-in-iis-server
* Use Service Account Keys !
  * https://stackoverflow.com/questions/57863574/google-sheets-api-v4-not-working-in-iis-server
  * https://cloud.google.com/iam/docs/creating-managing-service-account-keys