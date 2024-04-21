# Visual Pinball Automation Script

This PowerShell script automates various tasks related to managing ROMs, tables, and color files for Visual Pinball.

Just download the table, b2s, rom, and altColor rom (optional) to a clean downloads folder and run this. It will:

1. Prompt you to move the rom files
2. Extract files from downloaded zips
3. Prompt you to rename your .vpx files
4. Rename the .directb2s file to match
5. Move them to the tables folder
6. Create a folder in the altColor folder to prepare for any color DMD files
7. Prompt you to move color DMD files to their correct new folder
8. Auto-verify that the rom is named correctly
9. Update the registry for the rom to turn on the color dmd

## Features (as written by ChatGPT):

1. **Moving ROM Files:**
   - Identifies ZIP files in a specified source folder and allows the user to select and move ROM files from these ZIPs to a designated ROM folder.

2. **Extracting ZIP Files:**
   - Extracts the contents of remaining ZIP files in the source folder, renames certain files, and moves them to the source folder.

3. **Handling VPX and DirectB2S Files:**
   - Renames Visual Pinball tables (.vpx) and their corresponding DirectB2S backglass files based on user input and moves them to the tables folder.

4. **Managing AltColor Files:**
   - Creates folders in the AltColor directory based on ROM names and prompts the user to categorize color files into these folders.

5. **Standardizing AltColor File Names:**
   - Standardizes the names of color files within AltColor folders, ensuring they adhere to conventions.

6. **Cleanup:**
   - Offers an option to clean up the source folder by deleting all files.

The script guides the user through each step and provides relevant prompts and feedback along the way.
