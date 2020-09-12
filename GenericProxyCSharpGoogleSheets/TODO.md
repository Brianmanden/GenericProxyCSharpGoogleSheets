# Add class ConfigObj defining an object used to connect to SheetsService
## Should contain:
* SheetId
* Range
* DataOption (optional)
* InputOption (optional)

# CRUD
* ✓ Define the 4 CRUD methods as stubs
* Add functionality
* ✓ Create
* ✓ Read
* Update - https://stackoverflow.com/questions/37462887/update-a-cell-with-c-sharp-and-sheets-api-v4
* Delete

# Refactoring
* Move code into functions
* Tidy up main execution flow
* Write credentials.json to other location better suited for project so it can run on webserver
  * Can obj/Debug be used ?
  * https://stackoverflow.com/questions/57863574/google-sheets-api-v4-not-working-in-iis-server
* Use Service Account Keys !
  * https://stackoverflow.com/questions/57863574/google-sheets-api-v4-not-working-in-iis-server
  * https://cloud.google.com/iam/docs/creating-managing-service-account-keys