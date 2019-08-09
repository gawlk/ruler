from datetime import datetime
import googlemaps
import os.path
import re
import xlwt
import yaml

# ---
# Init
# ---

book = xlwt.Workbook(encoding="utf-8")
sheet = book.add_sheet("Kilometers")

file_name_key = "key.txt"
file_name_addresses = "addresses.yaml"
file_name_distances = "distances.yaml"
file_name_data = "data.txt"

# ---
# Google maps
# ---

file_name = file_name_key

try:
    print("Read key file...")
    key = open(file_name, "r").read()
except:
    print("ERROR: " + file_name + " file (for Google maps API) doesn't exist")
    print("Please create the file and paste the your key in it")
    quit()

try:
    print("Init Google maps...")
    gmaps = googlemaps.Client(key=key)
except:
    print("ERROR: The key for Google maps API is invalid")
    quit()

# ---
# Address book
# ---

file_name = file_name_addresses

try:
    print("Read addresses file...")
    f = open(file_name, "r", encoding="utf8")
except:
    print("ERROR: " + file_name + " file (for Google maps API) doesn't exist")
    quit()

try:
    print("Import addresses...")
    addresses_book = yaml.safe_load(f) or {}
except yaml.YAMLError as error:
    print("ERROR: " + file_name + " is incorrect")
    print(error)

# ---
# Distances book
# ---

file_name = file_name_distances

if os.path.isfile(file_name):
    print("Read the distances file...")

    f = open(file_name, "r", encoding="utf8")

    try:
        print("Import the distances...")
        distances_book = yaml.safe_load(f) or {}
    except yaml.YAMLError as error:
        print("ERROR: " + file_name + " is incorrect")
        print(error)
        quit()

# ---
# Data file
# ---

regex_date = re.compile("^(0|1|2|3)?[0-9](\/(0|1)?[0-9](\/(20)?(0|1|2)[0-9])?)?$")
regex_home = re.compile("^home ?= ?([a-z|A-Z|0-9]+_?)*[a-z|A-Z|0-9]")
regex_trip = re.compile("^(([a-z|A-Z|0-9]+_?)*[a-z|A-Z|0-9])? ?-?> ?(([a-z|A-Z|0-9]+_?)*[a-z|A-Z|0-9])?$")
regex_skip = re.compile("^(#.*)?$")

counter_global = 0
counter_trip = 0

day = None
month = None
year = None

home_name = ""
departure_name = ""
arrival_name = ""

file_name = file_name_data

try:
    print("Read addresses file...")
    f = open(file_name, "r", encoding="utf8")
except:
    print("ERROR: " + file_name + " file (for Google maps API) doesn't exist")
    quit()

print("Process the data file...")

for line in f:
    line = line.replace(" ", "").rstrip()
    counter_global += 1

    if (regex_date.match(line)):
        date = line.split("/")

        day = date[0]

        if (len(date) >= 2):
            month = date[1]

            if (len(date) == 3):
                year = date[2]
            else:
                print("ERROR: Incorrect date in " + file_name)
                quit()

        departure_name = ""
        arrival_name = ""

    elif (regex_home.match(line)):
        home_name = line.split("=")[1].strip()
        
        if (home_name not in addresses_book):
            print("ERROR: Unknown address for " + home_name + " in " + file_name + " (line " + str(counter_global) + ")")
            quit()
    elif (regex_trip.match(line)):
        counter_trip += 1
        trip = line.split(">")

        if day is None or month is None or year is None:
            print("ERROR: You must first initialize a complete date (ex: 31/12/1970) in " + file_name)
            quit()

        if trip[0]:
            departure_name = trip[0]
        elif not departure_name and home_name:
            departure_name = home_name
        elif not departure_name:
            print("ERROR: Incorrect departure in " + file_name + " (line " + str(counter_global) + ")")
            quit()

        if trip[1]:
            arrival_name = trip[1]
        elif home_name:
            arrival_name = home_name
        else:
            print("ERROR: Incorrect arrival in " + file_name + " (line " + str(counter_global) + ")")
            quit()

        if departure_name == arrival_name:
            print("ERROR: The departure and the arrival cannot be the same in " + file_name + " (line " + str(counter_global) + ")")
            quit()

        trip = departure_name + " > " + arrival_name

        if (departure_name not in addresses_book):
            print("ERROR: Address of " + departure_name + " is unknown in " + file_name + " (line " + str(counter_global) + ")")
            quit()
        elif (arrival_name not in addresses_book):
            print("ERROR: Address of " + arrival_name + " is unknown in " + file_name + " (line " + str(counter_global) + ")")
            quit()
        else:
            departure_address = addresses_book[departure_name]
            arrival_address = addresses_book[arrival_name]

            if (not distances_book or trip not in distances_book):
                print("Calculation of " + trip + " using Google Maps API...")

                results = gmaps.directions(
                    departure_address,
                    arrival_address,
                    mode = "driving",
                    departure_time = datetime.now().replace(hour=23),
                    alternatives = "true"
                )

                if (len(results) == 0):
                    print("ERROR: Google Maps couldn't calcute the distance of " + trip + " in " + file_name + " (line " + str(counter_global) + ")")
                    quit()

                distance = None

                for result in results:
                    result = result['legs'][0]['distance']['text'].split(" ")
                    new_distance = float(result[0])
                    unit = result[1]

                    if (unit == "m"):
                        new_distance = new_distance / 100

                    if (distance is None or new_distance < distance):
                        distance = new_distance
                
                distances_book[trip] = distance

            else:
                distance = distances_book[trip]

            sheet.write(counter_trip, 0, day + "/" + month + "/" + year)
            sheet.write(counter_trip, 1, departure_address)
            sheet.write(counter_trip, 2, arrival_address)
            sheet.write(counter_trip, 3, line.replace("_", " ").replace(">", " > "))
            sheet.write(counter_trip, 4, distance)

            departure_name = arrival_name
    elif (not regex_skip.match(line)):
        print("ERROR: Unknown line format in " + file_name + " (line " + str(counter_global) + ")")
        quit()

f.close()

# ---
# Save calculated distances
# ---

f = open(file_name_distances, "w")

if (distances_book):
    for key, value in distances_book.items():
        f.write(key + ": " + str(value) + "\n")

f.close()

print("Success !")
