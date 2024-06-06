# Matt Gurgiolo
# March 2024
# 
# Coot Sorting Algorithm

import pandas as pd
import numpy as np
import random as rd

class Sorter:
    
    def __init__(self, filepath):
        
        # Initialize Variables
        self.filepath = filepath
        self.student_data = None
        self.trip_data = None
        self.num_students = None
        self.num_trips = None
        self.num_assigned = None
        self.avg_preference = None
        self.categories = []
        self.statistics = None
        
        self.read_data(self.filepath)

    def read_data(self, filepath):
        
        # Read in the files
        self.student_data = pd.read_excel(filepath, sheet_name='Students')
        self.trip_data = pd.read_excel(filepath, sheet_name='Trips')
        
        # Initialize Variables
        self.num_students = self.student_data.shape[0]
        self.num_assigned = 0
        self.num_trips = self.trip_data.shape[0]
        self.avg_preference = 100
        self.categories = list(dict.fromkeys(trip for trip in self.trip_data['Category']))
        
        # Format Student Data
        self.student_data['Preferences'] = self.student_data.apply(lambda row: sorted(self.categories, key=lambda col: row[col]), axis=1)
        self.student_data = self.student_data.drop(columns=self.categories)
        self.student_data['Assigned'] = 0
        self.student_data['Preference Received'] = None

        # Format Trip Data
        self.trip_data['Gender'] = 0
        self.trip_data['Dorms'] = self.trip_data.apply(lambda row: [], axis=1)
        self.trip_data['Teams'] = self.trip_data.apply(lambda row: [], axis=1)
        self.trip_data['Students'] = self.trip_data.apply(lambda row: [], axis=1)
        # self.trip_data['Valid'] = 0
      
    def sort(self):
        
        # Pull out all unnassigned students, randomize, and loop through and place them if possible
        unassigned = self.student_data[self.student_data['Assigned'] == 0].sample(frac=1).reset_index(drop=True)
        for i, student in unassigned.iterrows():
            student_id = student['Student ID']  # Fetch the student's ID

            # Find a placement for the student using the modified function
            placement_trip_id = self.find_placement(student_id)

            # Assign student if a placement is found
            if placement_trip_id is not None:
                # Use the modified place function that works with IDs
                self.place(student_id, placement_trip_id)
                
    def wtsort(self):
        
        # Pull out all unnassigned students with water needs and randomize
        unassigned = self.student_data[self.student_data['Assigned'] == 0]
        unassigned = unassigned[unassigned['Water'] < 3].sample(frac=1).reset_index(drop=True)
        
        # Sort Water Needs
        for i, student in unassigned.iterrows():
            student_id = student['Student ID']  # Fetch the student's ID

            # Find a placement for the student using the modified function
            placement_trip_id = self.find_placement(student_id)

            # Assign student if a placement is found
            if placement_trip_id is not None:
                # Use the modified place function that works with IDs
                self.place(student_id, placement_trip_id)        
        
        # Pull out all unnassigned students with tent needs and randomize
        unassigned = self.student_data[self.student_data['Assigned'] == 0]
        unassigned = unassigned[unassigned['Tent'] < 3].sample(frac=1).reset_index(drop=True)
        
        # Sort Tent Needs
        for i, student in unassigned.iterrows():
            student_id = student['Student ID']  # Fetch the student's ID

            # Find a placement for the student using the modified function
            placement_trip_id = self.find_placement(student_id)

            # Assign student if a placement is found
            if placement_trip_id is not None:
                # Use the modified place function that works with IDs
                self.place(student_id, placement_trip_id)
                
        # Pull out all unnassigned students and randomize
        unassigned = self.student_data[self.student_data['Assigned'] == 0].sample(frac=1).reset_index(drop=True)
        
        # Sort the rest
        for i, student in unassigned.iterrows():
            student_id = student['Student ID']  # Fetch the student's ID

            # Find a placement for the student using the modified function
            placement_trip_id = self.find_placement(student_id)

            # Assign student if a placement is found
            if placement_trip_id is not None:
                # Use the modified place function that works with IDs
                self.place(student_id, placement_trip_id)
                

    def find_placement(self, student_id):
        
        # Find student row and check if the student is already assigned
        student = self.student_data[self.student_data['Student ID'] == student_id]
        if student['Assigned'].values[0] == 1:
            print(f"How'd I get here?!? I've already been assigned! {student_id}")
            return None

        # Loop through student Preferences
        stu_pref = list(student['Preferences'])
        preferences = stu_pref[0]
        # print(preferences)
        for trip_type in preferences:

            # Filter Trips for that preference and loop through those trips
            filtered_trips = self.trip_data[self.trip_data['Category'] == trip_type].sort_values(by='Capacity', ascending=False).sample(frac=1).reset_index(drop=True)
            
            for _, trip in filtered_trips.iterrows():

                # Check compatibility
                if student['Water'].values[0] < 3 and trip['Water']:
                    continue
                elif student['Tent'].values[0] < 3 and trip['Tent']:
                    continue
                elif student['Dorm'].values[0] in trip['Dorms']:
                    continue
                elif pd.notna(student['Team'].values[0]) and student['Team'].values[0] in trip['Teams']:
                    continue
                elif abs(trip['Gender'] + student['Gender'].values[0]) > 3:
                    continue
                elif trip['Capacity'] == 0:
                    continue
                else:
                    self.student_data.loc[self.student_data['Student ID'] == student_id, 'Preference Received'] = preferences.index(trip_type) + 1
                    return trip['Trip']
    
    
    def place(self, student_id, trip_id):
        
        # Fetch the trip row using trip_id
        trip_row = self.trip_data[self.trip_data['Trip'] == trip_id]

        # Assuming there's only one matching row for each trip_id
        if not trip_row.empty:
            trip_index = trip_row.index[0]
            # Update trip parameters
            self.trip_data.at[trip_index, 'Capacity'] -= 1
            self.trip_data.at[trip_index, 'Gender'] += self.student_data.loc[self.student_data['Student ID'] == student_id, 'Gender'].values[0]
            self.trip_data.at[trip_index, 'Dorms'].append(self.student_data.loc[self.student_data['Student ID'] == student_id, 'Dorm'].values[0])
            self.trip_data.at[trip_index, 'Teams'].append(self.student_data.loc[self.student_data['Student ID'] == student_id, 'Team'].values[0])
            self.trip_data.at[trip_index, 'Students'].append(student_id)

            # Mark student as assigned
            self.student_data.loc[self.student_data['Student ID'] == student_id, 'Assigned'] = 1
            self.num_assigned += 1
       
            
    def remove(self, student_id, trip_id):
        
        # Fetch the trip row using trip_id
        trip_row = self.trip_data[self.trip_data['Trip'] == trip_id]

        # Assuming there's only one matching row for each trip_id
        if not trip_row.empty:
            trip_index = trip_row.index[0]
            # Update trip parameters
            self.trip_data.at[trip_index, 'Capacity'] += 1
            self.trip_data.at[trip_index, 'Gender'] -= self.student_data.loc[self.student_data['Student ID'] == student_id, 'Gender'].values[0]
            self.trip_data.at[trip_index, 'Dorms'].remove(self.student_data.loc[self.student_data['Student ID'] == student_id, 'Dorm'].values[0])
            self.trip_data.at[trip_index, 'Teams'].remove(self.student_data.loc[self.student_data['Student ID'] == student_id, 'Team'].values[0])
            self.trip_data.at[trip_index, 'Students'].remove(student_id)

            # Mark student as unassigned
            self.student_data.loc[self.student_data['Student ID'] == student_id, 'Assigned'] = 0
            self.num_assigned -= 1
      
                
    def find_trip_to_remove(self, student_id):
        
        # Find student row and check if the student is already assigned
        student = self.student_data[self.student_data['Student ID'] == student_id]
        if student['Assigned'].values[0] == 1:
            print(f"How'd I get here?!? I've already been assigned! {student_id}")
            return None

        # Loop through student Preferences
        preferences = list(student['Preferences'])
        for trip_type in preferences[0]:

            # Filter Trips for that preference and loop through those trips
            filtered_trips = self.trip_data[self.trip_data['Category'] == trip_type]
            for _, trip in filtered_trips.iterrows():

                # Check compatibility
                if student['Water'].values[0] < 3 and trip['Water']:
                    continue
                elif student['Tent'].values[0] < 3 and trip['Tent']:
                    continue
                elif student['Dorm'].values[0] in trip['Dorms']:
                    continue
                elif pd.notna(student['Team'].values[0]) and student['Team'].values[0] in trip['Teams']:
                    continue
                elif abs(trip['Gender'] + student['Gender'].values[0]) > 3:
                    continue
                elif trip['Capacity'] == 0:
                    continue
                else:
                    return trip['Trip']
    
    
    def generate_statistics(self):
        
        # Calculate the percentage of students assigned to trips
        percent_assigned = (self.num_assigned / self.num_students) * 100 if self.num_students > 0 else 0
        self.avg_preference = self.student_data['Preference Received'].mean()

        # Create a statistics DataFrame
        stats_data = {
            'Total Students': [self.num_students],
            'Number Assigned': [self.num_assigned],
            'Average Preference': [self.avg_preference],
            'Percent Assigned': [percent_assigned],
        }
        self.statistics = pd.DataFrame(stats_data)
    
    def write(self, filepath):
        
        self.generate_statistics()
        
        # Write the data to a file
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            self.student_data.to_excel(writer, sheet_name='Students', index=False)
            self.trip_data.to_excel(writer, sheet_name='Trips', index=False)
            self.statistics.to_excel(writer, sheet_name='Statistics', index=False)
            
    def run(self):
        
        # This runs the first pass of sorting
        self.wtsort()
        
        
    # Implement further Sorting
        # Run back through unsorted students? 
        # Check Valid Trips