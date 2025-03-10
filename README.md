# Python NBT Data Parser + Extractor (nbtdatparse.py)
![image](https://github.com/user-attachments/assets/d78feab2-db47-4c15-9def-96c5117d0a8f)
This is specifically for scanning many minecraft .dat files at once and listing them on an excel sheet to more easily find old or deleted minecraft seeds. My use case is finding the seed to my minecraft world I deleted 13 years ago. (.dat and log.gz file have the same identifier so I may add an option to search recovered log files for /seed or for specific usernames or info)

- brief instructions for [DMDE .dat searching](optionalDMDE/info.md)
(free software that can recover up to 4000 files, one directory at a time)

***nbtdatparse.py retrieves:***
- file name
- seed
- playtime
- generator used
- level name
- game mode
- spawn location
- file path
- (possible to get ai to augment code to also include data from the inventory or [other information that is typically stored in the level.dat files](https://minecraft.wiki/w/Java_Edition_level_format))

## 0. Download my .py file
place file in directory to scan


## 1. Install Python / use a working interpreter version
worked for me:
- 3.8.12
- 3.11 ?

didn't work for me:
- 3.12.9



## 2. Install Required Libraries
- nbtlib
   ```
   pip install nbtlib
   ```
- openpyxl
   ```
   pip install openpyxl
   ```

## 3. Set scan and save locations in the .py file (optional)
By default it scans all .dat files in the current folder and subdirectories, then it saves `minecraft_worlds.xlsx` in the same folder.

## 4. Should be good to go!
I used Perplexity to create this script and it may help someone. I'll put my short quide for [DMDE](optionalDMDE/info.md) to scan whole drives for data files and my scan signatures.

## 5. Sorting the .xlsx
1. Highlight the top row of the Data tab
- Click sort and filter
- Select filter
2. The top row should now have dropdown arrows, select and sort however you need

----
## here is what mine looked like at the end
![2025-03-10 all is well](https://github.com/user-attachments/assets/52eaa2e9-7e88-49b3-9d7b-da7ebb16b6f3)
errors are to be expected, I have it set to recover more data than less because it is easy to sort through

