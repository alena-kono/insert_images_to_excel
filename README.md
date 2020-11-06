# Insert images to excel


Python script that inserts images to excel (.xls*) file.

## Motivation

This is a completely educational project

## Build Status

- [ ] TODO display list with images that not found
- [ ] TODO Set property to all images: move and size with cells
- [ ] TODO Add compression to the images
- [ ] TODO Add axception handling

## Installation

To launch this bot run run.py inside "insert_images_to_excel/insert_images_to_excel" folder

OS X & Linux:

```sh
pipenv install
```
```sh
cd insert_images_to_excel
python3 run.py
```

# Utilization

You will need .xls or .xlsx file that contains:
* at least one target column (column to where images will be inserted);
* column with names exactly (excluding extension) matching your images

![picture alt](/Users/alena/Dev/heart_repos/orbico/insert_images_to_excel/test_data/target_file.png "target file")
