# Folder Cleanup

This process was designed as an automation to address an issue the business group was facing. Copies of appeals and files are saved to a share drive and must be kept for a minimum of 7 years. Staff continually need to search and save to this folder, both of which were taking excessively long to execute due to the size of the folder (nearly 600 GB). This process is run weekly and performs the following:

1. Unpacks subdirectories and moves files to the main directory to improve search time
2. Moves files to an archive folder that do not contain a valid invoice number (ie will be unsearchable in the future)
3. Moves files to an archive folder that were created over a year ago
4. Zips files over 500mb and moves them to an archive folder

The archive folder is organized by year and month of the creation date so if they ever need to be searched, the user can navigate directly to the month and year the file should have been saved to search.

This was executed in a Jupyter notebook so each individual section could be run separately based on need. Variables are configurable to change source and destination locations, file size requirements, file age requirements, and valid/invalid extension types.
