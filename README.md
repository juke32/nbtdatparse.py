# Python NBT Data Parser + Extractor (nbtdatparse.py)
![image](https://github.com/user-attachments/assets/d78feab2-db47-4c15-9def-96c5117d0a8f)
This is specifically for scanning many minecraft .dat files at once and listing them to be located, or for relevant data to be located. My use case is finding the seed to my minecraft world I deleted 13 years ago.

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
can be seperate from your .dat files


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

## 3. Set scan and save locations in the .py file
By default it scans all of the subdirectories of `D:\dump`(top of file) and then saves the results to `minecraft_worlds.xlsx`in to `D:\dump`(bottom of file)

## 4. Should be good to go!
I used Perplexity to create this script and it may help someone. I'll put my short quide for [DMDE](optionalDMDE/info.md) to scan whole drives for data files and my scan signatures.

## 5. Sorting the .xlsx
1. Highlight the top row of the Data tab
- Click sort and filter
- Select filter
2. At the top of the page select the dropdown arrows and select and sort however you need

----
## here is what mine looked like at the end
![2025-03-10 all is well](https://github.com/user-attachments/assets/52eaa2e9-7e88-49b3-9d7b-da7ebb16b6f3)
errors are to be expected, I have it set to recover more data than less because it is easy to sort through

