# Python NBT Data and Log Parser + Extractor (nbtdatparse.py)
![image](https://github.com/user-attachments/assets/99248ac3-c377-4b00-b10a-469248564737)
This is specifically for scanning many minecraft .dat files at once and listing them on an excel sheet to more easily find old or deleted minecraft seeds. My use case is finding the seed to my minecraft world I deleted 13 years ago.

- Short guide for [DMDE](optionalDMDE/info.md) minecraft .dat or log file recovery using my scan signatures.
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
- searches logs for instances of 'seed'
- (possible to get ai to augment code to also include data from the inventory or [other information that is typically stored in the level.dat files](https://minecraft.wiki/w/Java_Edition_level_format))


## 0. Download my .py file
place file in good location


## 1. Install Python
3.8 and 3.11 worked for me, each on a seperate machine



## 2. Install Required Libraries
install the libraries to the interpreter / version of python that runs the code
- nbtlib
   ```
   pip install nbtlib
   ```
- openpyxl
   ```
   pip install openpyxl
   ```

## 3. Set `directory_path` for scanning and output of `minecraft_worlds_recovery.xlsx` - not optional
- ~~By default it scans all .dat files in the current folder and subdirectories, then it saves `minecraft_worlds.xlsx` in the same folder.~~
- current examples: `D:/dump` & `C:/Users/juke32/AppData/Roaming/.minecraft/saves`  


## 4. Should be good to run!
If it doesn't work double check the file path, direction of the slashes, if the correct python interpreter is used, try using a terminal window not an ide or coding enviroment.

## 5. Sorting through the .xlsx
1. Highlight the top title row of the data in any tab
- Click sort and filter
- Select filter
2. The top row should now have dropdown arrows, select and sort however you need
3. If you uncheck errors/seeds/info you already know about, you can reduce the clutter I didn't code out

