> **Warning**<br>
> This documentation is under completion.

### Table of Contents
&emsp;&emsp;[A. Installation](#section-a-installation)

## Section A: Installation

### Installation
- Download **mic-drop-results.zip** file of the latest version from the following page<br>https://github.com/berkeleyfx/mic-drop-results/releases
- Extract the ZIP file
- This new folder will be your working directory<br>![image](https://user-images.githubusercontent.com/106049382/195757100-d220565d-360f-460b-920a-5754877219bd.png)
- After that, download the template files (**data.xlsx**, **template.pptm**, and **token.txt**) from [`/sample`](./sample) and put the files in your working directory<br>![image](https://user-images.githubusercontent.com/106049382/195757406-5fb450db-f959-4219-abf4-989b54d7831f.png)

### Getting your tokens
- Tokens are used to access Discord API to download profile pictures
- You can have more than one token in **token.txt** to avoid getting rate-limited
- You can follow these steps to get your own tokens<br>https://mee6.xyz/tutorials/how-to-generate-a-custom-bot-token
- You do not have to invite the bot to your server to get it working
- Because an application can only have one bot, you need to create multiple applications to create multiple bots, and therefore, multiple tokens.

## Section B: Get the Program Running

### Customizing your templates
- Every slide in **template.pptm** is called a template
- You can assign a template to a contestant through the `template` column of **data.xlsx**
- More advanced customization at the final section (helpful for the designing process at the beginning of a season or for special rounds)
- There is also a contestant database, which stores the user IDs of all the contestants, at the final section. Collecting user IDs only needs to be done once every season.

### Running the program
- Simply run **app.exe**

### Debugging methods
> **Note**<br>
> You may have been redirected here as you encountered an error while running the program.<br><br>
> You can disable this feature in **settings.json** by setting `debug_methods` to 0.<br>
> Setting this to 1 will only disable redirecting for soft-handled error.

- If you encounter any bugs, please follow these steps:

  1. Check **data.xlsx** very carefully. Make sure:
      - none of the columns are blank 
      - there is no unnecessary sheet in the workbook
      - all cells around the data regions are blank **(this is a common mistake)**
  2. If the problem still persists, please report the bug to me on Discord. [(Banz#6175)](https://discord.com/users/1010885414850154587)
  3. If you believe an error should be soft-handled (automatically fill N/A values
