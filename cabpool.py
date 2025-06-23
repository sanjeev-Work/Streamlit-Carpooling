# -*- coding: utf-8 -*-
"""
This script assigns passengers to cabs for hotel, support, and airport rides
and outputs results to an Excel workbook. It enforces a maximum of 3 travelers
per cab. Within each predefined group (as returned by Parser.ride_to_*), it
tries to keep people using the same `app` together whenever possible, but never
at the expense of increasing the total number of cabs. Cab identifiers use:
- Two-digit numbers for hotel rides ('01', '02', ...)
- Three-digit numbers for support rides ('100', '101', ...)
- Letters for airport rides ('A', 'B', ..., 'Z', 'AA', 'BB', ...)
No cab with exactly one occupant will be recorded—those riders get a blank entry
unless they are marked as a personal ride. Each non-personal cab ID is prefixed
with "Cab ".
"""

import sys
from collections import defaultdict
from givetochat1_6 import Parser, Person
from openpyxl import Workbook
from openpyxl.styles import Font

# Global counters to ensure uniqueness of cab IDs across all groups per segment
_next_hotel_cab = 0
_next_support_cab = 0
_next_airport_cab = 0


def _generate_airport_cab_id(index: int) -> str:
    """
    Generate cab IDs for airport rides:
    'A', 'B', ..., 'Z', 'AA', 'BB', ..., 'ZZ', etc.
    After 'Z', repeats of the same letter in pairs.
    """
    length = (index // 26) + 1
    letter = chr(ord('A') + (index % 26))
    return letter * length


def _assign_with_app_preference(people: list) -> list:
    """
    Given a list of Person objects (all in the same predefined group),
    split them into sub-cabs of size <= 3, minimizing total cabs and
    keeping same-app riders together whenever possible.

    Returns a list of cabs, where each cab is a list of Person.
    """
    n = len(people)
    if n == 0:
        return []
    if n <= 3:
        return [people.copy()]

    # 1. Bucket people by app
    app_buckets = defaultdict(list)
    for p in people:
        app_buckets[p.app].append(p)

    cab_units = []
    leftovers_two = []   # 2-person lists
    leftovers_one = []   # single Person

    # 2. Extract full triples of same-app
    for app, bucket in list(app_buckets.items()):
        while len(bucket) >= 3:
            trio = [bucket.pop() for _ in range(3)]
            cab_units.append(trio)
        if len(bucket) == 2:
            pair = [bucket.pop(), bucket.pop()]
            leftovers_two.append(pair)
        if len(bucket) == 1:
            single = bucket.pop()
            leftovers_one.append(single)

    # 3. Pair 2-person leftovers with a single-person leftover if possible
    for pair in leftovers_two[:]:
        if leftovers_one:
            single = leftovers_one.pop()
            cab_units.append(pair + [single])
            leftovers_two.remove(pair)
    for pair in leftovers_two:
        cab_units.append(pair)
    leftovers_two.clear()

    # 4. Group remaining singles (across apps) into cabs of up to 3
    idx = 0
    while idx < len(leftovers_one):
        chunk = leftovers_one[idx : idx + 3]
        cab_units.append(chunk.copy())
        idx += 3

    return cab_units


def _assign_cabs(groups: dict, id_type: str, max_per_cab: int = 3) -> dict:
    """
    Assign people in each group (groups[key] is a list of Person) to cabs,
    splitting groups across multiple cabs if they exceed max_per_cab riders.
    Uses the same-app-preference algorithm internally and ensures no cab ID
    is reused within the given segment type. Does not allow more than max_per_cab
    per cab.

    Returns:
        dict: Mapping from group key to a dict {cab_id: [Person, ...], ...}.
    """
    global _next_hotel_cab, _next_support_cab, _next_airport_cab
    assignments = {}

    for key, people in groups.items():
        if not people:
            assignments[key] = {}
            continue

        # Split the group into sub-cabs of size <= 3 with app preference
        sub_cabs = _assign_with_app_preference(people)

        cab_map = {}
        for riders in sub_cabs:
            # skip sub-cabs with just one person
            if len(riders) < 2:
                continue # and don't generate an index
            if id_type == 'hotel':
                idx = _next_hotel_cab
                cab_id = f"{idx + 1:02d}"
                _next_hotel_cab += 1
            elif id_type == 'support':
                idx = _next_support_cab
                cab_id = f"{100 + idx:03d}"
                _next_support_cab += 1
            elif id_type == 'airport':
                idx = _next_airport_cab
                cab_id = _generate_airport_cab_id(idx)
                _next_airport_cab += 1
            else:
                raise ValueError(f"Unknown id_type: {id_type}")

            cab_map[cab_id] = riders

        assignments[key] = cab_map

    return assignments


def write_cab_excel(input_xlsx: str, output_xlsx: str):
    # 1. Parse input
    all_people, h_personal, a_personal = Parser.process_excel(input_xlsx)

    # 2. Build segment groupings
    # Hotel: only those without personal Hotel flag
    hotel_candidates = [p for p in all_people if p not in h_personal]
    ride_hotel = Parser.ride_to_hotel(hotel_candidates, thresh=0) # strict flight grouping

    # Support: assume everyone needs a cab (no personal override for support)
    ride_support = Parser.ride_to_support(all_people)

    # Airport: only those without personal Airport flag
    airport_candidates = [p for p in all_people if p not in a_personal]
    ride_airport = Parser.ride_to_airport(airport_candidates)

    # 3. Assign to cabs (maximum 3 riders per cab), using global counters
    hotel_cab_assignments = _assign_cabs(ride_hotel, id_type='hotel', max_per_cab=3)
    support_cab_assignments = _assign_cabs(ride_support, id_type='support', max_per_cab=3)
    airport_cab_assignments = _assign_cabs(ride_airport, id_type='airport', max_per_cab=3)

    # Build lookup: person_name -> assigned cab ID per ride type, skipping solo cabs
    assigned = {p.name: {} for p in all_people}

    # Hotel assignments
    for cab_map in hotel_cab_assignments.values():
        for cab_id, riders in cab_map.items():
            if len(riders) > 1:
                for p in riders:
                    assigned[p.name]['hotel'] = cab_id

    # Support assignments
    for cab_map in support_cab_assignments.values():
        for cab_id, riders in cab_map.items():
            if len(riders) > 1:
                for p in riders:
                    assigned[p.name]['support'] = cab_id

    # Airport assignments
    for cab_map in airport_cab_assignments.values():
        for cab_id, riders in cab_map.items():
            if len(riders) > 1:
                for p in riders:
                    assigned[p.name]['airport'] = cab_id

    # 4. Create workbook and sheet
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = 'Cab Assignments'
    ws.append(['Name', 'Cab to Hotel', 'Cab to Support', 'Cab to Airport'])

    for person in sorted(all_people, key=lambda p: p.name):
        name = person.name

        # Cab to Hotel: personal override or assigned cab (skip solo)
        if person.personal['Hotel']:
            hotel_cab = 'Personal'
        else:
            hotel_cab = assigned[name].get('hotel', '')
            if hotel_cab:
                hotel_cab = f"Cab {hotel_cab}"

        # Cab to Support: no personal override; only if assigned with >1 riders
        support_cab = assigned[name].get('support', '')
        if support_cab:
            support_cab = f"Cab {support_cab}"

        # Cab to Airport: personal override or assigned cab (skip solo)
        if person.personal['Airport']:
            airport_cab = 'Personal'
        else:
            airport_cab = assigned[name].get('airport', '')
            if airport_cab:
                airport_cab = f"Cab {airport_cab}"

        ws.append([name, hotel_cab, support_cab, airport_cab])

    wb.save(output_xlsx)
    print(f"Wrote cab assignments to {output_xlsx}")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python cabpool.py input.xlsx output.xlsx")
        sys.exit(1)
    write_cab_excel(sys.argv[1], sys.argv[2])
