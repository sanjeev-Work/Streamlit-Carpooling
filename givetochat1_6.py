import pandas as pd
from datetime import datetime, timedelta, time
from collections import defaultdict
import re

class Person:
    def __init__(self, row):
        # Initialize person attributes based on row data from Excel file
        self.name = row.get("Name", "NoName")
        self.hotel = row.get("Hotel", "NoHotel")
        if(pd.isna(self.hotel)): self.hotel = "NoHotel"

        self.app = row.get("App", "NoApp")
        if pd.isna(self.app):
            self.app = "NoApp"

        # Determine role (CC or FS), defaulting to CC if missing
        self.role = row.get("CC or FS", "CC")
        if pd.isna(self.role) or self.role.strip() == "": self.role = "CC"
        self.location = row.get("Location", "NoLocation")  # Store location info
        if pd.isna(self.location) or self.location.strip() == "": self.location = "NoLocation"

        # Check if person already has a rental car assigned
        self.has_rental_car = row.get("Rental Car") == self.name

        # Parse date and time fields safely
        self.begin_onsite = self.safe_parse_datetime(row.get("Begin OnSite"))
        self.end_onsite = self.safe_parse_datetime(row.get("End OnSite"))
        self.depart_date = self.safe_parse_datetime(row.get("Depart Date"))
        self.return_date = self.safe_parse_datetime(row.get("Return Date"))

        # save whether or not this person can be assigned a rental car (pull from insert ETR info page)
        rc_etr = row.get("Rental Car_etr", "")
        self.can_drive = pd.isna(rc_etr) or (isinstance(rc_etr, str) and rc_etr.strip() == "")

        # track whether this person has been assigned a rental car at another point in the process
        # Automaticaly assign rental car if float + can drive
        if "float" in self.location.lower() and self.can_drive:
            self.given_rental_car = True
        else:
            self.given_rental_car = False

#        if self.can_drive and self.begin_onsite != datetime.min and self.end_onsite != datetime.min:
#            stay_duration = self.end_onsite - self.begin_onsite
#            if stay_duration >= timedelta(days=10):
#                self.given_rental_car = True # assign rental cars to travelers whose onsite duration is 10+ days # TODO
        
        # Determine personal travel flags for hotel and airport rides (case-insensitive lookup)
        hotel_personal = False
        airport_personal = False
        # Find actual column names ignoring case
        hotel_col = next((col for col in row.index if col.lower() == "ride to hotel"), None)
        if hotel_col:
            val = row.get(hotel_col, "")
            if isinstance(val, str) and val.strip().lower() == "personal":
                hotel_personal = True

        airport_col = next((col for col in row.index if col.lower() == "ride to airport"), None)
        if airport_col:
            val = row.get(airport_col, "")
            if isinstance(val, str) and val.strip().lower() == "personal":
                airport_personal = True
        self.personal = {"Hotel": hotel_personal, "Airport": airport_personal}
        
        
        
        # Parse flight information
        self.arrival_flight = self.parse_flight_info(row.get("Arrival Flight"))
        if self.arrival_flight["Time"] is not None:
            # store as datetime for easier difference computations
            self.arrival_dt = datetime.combine(self.depart_date.date(), self.arrival_flight["Time"])
        else:
            self.arrival_dt = None
        self.return_flight = self.parse_flight_info(row.get("Return flight"))
        if self.return_flight["Time"] is not None:
            self.return_dt = datetime.combine(datetime.today(), self.return_flight["Time"])
        else:
            self.return_dt = None
    

    @staticmethod
    def safe_parse_datetime(date_str):
        # Safely parse datetime from string format, return min date if invalid
        if pd.isna(date_str) or not isinstance(date_str, str):
            return datetime.min
        try:
            return datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            return datetime.min

    @staticmethod
    def parse_flight_info(flight_str):
        """Parses flight details including flight number, time, and city."""
        if not flight_str or not isinstance(flight_str, str):
            return {"Flight Number": "NoFlightNum", "Time": time.min, "City": "NoCity"}
        
        match = re.match(r"(\d+)@\s*([\d:]+)\s*([ap]?)\s*(.*)?", flight_str, re.IGNORECASE)

        if match:
            flight_number = match.group(1)
            time_str = match.group(2)
            am_pm = match.group(3).upper() if match.group(3) else ""
            city = match.group(4).strip() if match.group(4) else "NoCity"
            
            try:
                dt = datetime.strptime(f"{time_str}{am_pm}M", "%I:%M%p")
                time_obj = dt.time()
            except ValueError:
                time_obj = time.min

            return {"Flight Number": flight_number, "Time": time_obj, "City": city}
        else:
            return {"Flight Number": "NoFlightNum", "Time": time.min, "City": "NoCity"}
        
    def __str__(self):
        """Returns the person's name with spaces replaced by hyphens."""
        return self.name.replace(" ", "-")
    
    def __repr__(self):
        return self.__str__()
    
class Parser:

    @staticmethod
    def process_excel(input_file):
        """
        Reads an Excel file and processes Person data.
        returns: all people in input file, people marked personal for hotel, people marked personal for airport
        """
        df_main = pd.read_excel(input_file, sheet_name="Car pool", dtype=str, skiprows=1)
        df_etr = pd.read_excel(input_file, sheet_name="Insert ETR Info Here", usecols=["Name", "Rental Car"], dtype=str)

        df = df_main.merge(df_etr, on="Name", how="left", suffixes=("", "_etr"))
    
        # Create lists of Person objects from Excel data
        people = [Person(row) for _, row in df.iterrows()]
        h_personal = [p for p in people if p.personal["Hotel"]]
        a_personal = [p for p in people if p.personal["Airport"]]
    
        return people, h_personal, a_personal
    
    @staticmethod
    def list_hotels(all_people: list[Person]) -> set:
        """
        reads exact hotel names from already processed people list for GPT to reference when merging
        """
        hotels = set()
        for person in all_people:
            if person.hotel and isinstance(person.hotel, str):
                hotels.add(person.hotel)
        return hotels



    @staticmethod
    def merge_hotels(all_people, hotel_groups):
        """
        Merges specified hotels by updating each Person's hotel attribute.
        hotel_groups: list of lists, where each inner list contains hotel names to merge.
        After merging, the hotel's value becomes a combined name (e.g., "HotelA/HotelB").
        Example:
            Parser.merge_hotels(people, [["Hotel A", "Hotel B"], ["Hotel C", "Hotel D"]])
        """
        for group in hotel_groups:
            # Determine merged name by joining sorted names
            merged_name = "/".join(sorted([h.strip() for h in group]))
            for p in all_people:
                if p.hotel and p.hotel.strip() in group:
                    p.hotel = merged_name

    @staticmethod
    def _cluster_by_time(travelers,  attribute, threshold_hours=1.0):
        """
        Bottom-up complete linkage clustering on traveler arrival times.
        Groups clusters such that the maximum pairwise time difference in each cluster <= threshold_hours.
        Returns a list of cluster lists.
        """
        # Initialize each traveler as its own cluster
        clusters = [[p] for p in travelers]

        # Helper: extract timestamp floats (seconds) for clusters
        def times_list(cluster, attr=attribute):
            timestamps = []
            for p in cluster:
                dt = getattr(p, attr, None)
                if isinstance(dt, datetime):
                    timestamps.append(dt.timestamp())
            return timestamps

        # Distance between two clusters: max pairwise difference (complete linkage)
        def cluster_dist(c1, c2):
            times1 = times_list(c1)
            times2 = times_list(c2)
            if not times1 or not times2:
                return float('inf')
            return max(abs(t1 - t2) for t1 in times1 for t2 in times2)
        
        # Merge clusters while minimal distance <= threshold
        thresh = threshold_hours * 3600 # translate into seconds
        merged = True
        while merged:
            merged = False
            min_dist = float('inf')
            pair = (None, None)
            # find closest pair of clusters
            for i in range(len(clusters)):
                for j in range(i+1, len(clusters)):
                    d = cluster_dist(clusters[i], clusters[j])
                    if d < min_dist:
                        min_dist = d
                        pair = (i, j)
            # merge if below threshold
            if pair[0] is not None and min_dist <= thresh:
                i, j = pair
                clusters[i].extend(clusters[j])
                del clusters[j]
                merged = True
        return clusters
    
    @staticmethod
    def ride_to_hotel(all_people, thresh=0.5):
        """Organizes carpools from airport to hotel, ensuring efficient grouping."""
        if thresh == 0: 
            use_flight_number = True
        else:
            use_flight_number = False

        depart_hotel_groups = defaultdict(list) # grouped by city, depart date, hotel

        people = [p for p in all_people if not p.personal["Hotel"]] # don't group hotel personal travelers -- should already be filtered out

        # Step 0: sort input for cleaner debug output
        people.sort(key=lambda p: (p.hotel.strip(), p.depart_date if p.depart_date else datetime.min))

        # separate ID/IE/Exemplar/Emeritus to ensure no passengers
        remainder = []
        for p in people:
            if p.app in ("ID", "IE") or any(kw in p.name.lower() for kw in ("emeritus", "exemplar")):
                key = f"{p.app} | {p.name}"
                depart_hotel_groups[key].append(p)
            else:
                remainder.append(p)
        people = remainder # no more ID/IE people

        
        # Step 1: Group travelers by depart date & hotel & arrival flight city
        for person in people:
            key = f"{person.hotel.strip()} | {person.depart_date.date() if person.depart_date else 'Unknown'} | {person.arrival_flight['City']}"
            if use_flight_number:
                flight_key = person.arrival_flight.get("Flight Number", "NoFlightNum")
                key = f"{key} | {flight_key}"
            depart_hotel_groups[key].append(person)

        depart_hotel_flight_groups = defaultdict(list)

        # Step 2: further divide groups by clustering flight times
        for key, group in depart_hotel_groups.items():
            clusters = Parser._cluster_by_time(group, attribute="arrival_dt", threshold_hours=thresh)
            for cl in clusters:
                times = sorted([p.arrival_dt for p in cl if p.arrival_dt])
                start = times[0].strftime('%H:%M') if times else 'Unknown'
                end = times[-1].strftime('%H:%M') if times else 'Unknown'
                depart_hotel_flight_groups[f"{key} | {start}-{end}"] = cl

        return depart_hotel_flight_groups
    

    @staticmethod
    def ride_to_support(people):
        # Sort input list by FS vs CC, then App, then by hotel, then by support location, then by begin_onsite
        people.sort(key=lambda p: (p.role, p.app, p.hotel, p.location, p.begin_onsite if p.begin_onsite else datetime.min))
        
        role_hotel_groups = defaultdict(list)
        
        for person in people:
            begin_onsite_key = person.begin_onsite.strftime('%Y-%m-%d %H:%M') if person.begin_onsite else 'Unknown'
            end_onsite_key = person.end_onsite.strftime('%Y-%m-%d %H:%M') if person.end_onsite else 'Unknown'
            key = f"{person.role} | {person.hotel.strip()} | {person.location.strip()} | {begin_onsite_key} to {end_onsite_key}"
            # separate ID, IE, Emeritus, Exemplar
            if person.app in ("ID", "IE") or any(kw in person.name.lower() for kw in ("emeritus", "exemplar")):
                key = f"{person.name} | {key}"
            role_hotel_groups[key].append(person)
        
        return role_hotel_groups # only exact matches for all parameters
    
    @staticmethod
    def ride_to_airport(all_people, thresh=1.0):
        """Organizes carpools from hotel to airport based on return flight details."""
        if thresh == 0:
            use_flight_number = True
        else:
            use_flight_number = False

        groups_noTime = defaultdict(list)
        people = [p for p in all_people if not p.personal["Airport"]] # filter out personal
        cutoff = time(11, 0, 0) # 11am
        people = [p for p in people if p.end_onsite.time() >= cutoff] # filter out potential night shifters for manual review
        
        # Step 0: sort input for cleaner debug output
        people.sort(key=lambda p: (p.hotel.strip(), p.depart_date if p.depart_date else datetime.min))

        # separate ID/IE/Exemplar/Emeritus to ensure no passengers
        remainder = []
        for p in people:
            if p.app in ("ID", "IE") or any(kw in p.name.lower() for kw in ("emeritus", "exemplar")):
                key = f"{p.app} | {p.name}"
                groups_noTime[key].append(p)
            else:
                remainder.append(p)
        people = remainder # no more ID/IE people

        # separate leaving from location / leaving from hotel
        from_loc = [p for p in people if p.end_onsite.date() == p.return_date.date()] # end onsite == return date
        from_hotel = [p for p in people if p.end_onsite.date() != p.return_date.date()] # end onsite != return date

        for person in from_loc:
            key = f"{person.location} | {person.return_flight['City']} | {person.end_onsite.date()}"
            if use_flight_number:
                flight_key = person.return_flight.get("Flight Number", "NoFlightNum")
                key = f"{key} | {flight_key}"
            groups_noTime[key].append(person)

        for person in from_hotel:
            key = f"{person.hotel} | {person.return_flight['City']} | {person.end_onsite.date()}"
            groups_noTime[key].append(person)


        carpool_groups = defaultdict(list)

        for key, group in groups_noTime.items():
            clusters = Parser._cluster_by_time(group, attribute="return_dt", threshold_hours=thresh)
            for cl in clusters:
                times = sorted ([p.return_dt for p in cl if p.return_dt])
                start = times[0].strftime('%H:%M') if times else 'Unknown'
                end = times[-1].strftime('%H:%M') if times else 'Unknown'
                carpool_groups[f"{key} | {start}-{end}"] = cl
        
        for p in people:
            if p.end_onsite.time() >= cutoff:
                groups_noTime["Potential Night Shift"].append(p)
            

        return carpool_groups
        