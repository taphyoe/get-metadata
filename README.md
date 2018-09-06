# get-metadata
Powershell script to get file metadata and folder size. Gets all the metadata and returns a CSV file (with | separation).

-----------------------------------------------------------------------------
Script: Get-Metadata

Version: 2.1
Author: Tin Aung Phyoe (taphyoe@gmail.com)
Created Date: 05/09/2018
Keywords: Metadata, Storage, Files
Credits: Antonio Gatti
Last Modified: 06/09/2018 11:16 AM
comments: to get file metadata and folder size
Gets all the metadata and returns a CSV file (with | separation)

PARAMETERS:
-folder: folder to be scanned
-depth: set how deep you want the script to go and get the file metadata. After this level it will only return the folder size. (default value is 5)
-maxFileRows: set the max number of rows you want in the CSV. It will generate another file every time the script reach the limit.

-----------------------------------------------------------------------------
