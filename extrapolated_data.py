from openpyxl import Workbook, load_workbook
import random
from datetime import datetime,timedelta

# Create a new workbook and set the active sheet
wb = Workbook()
ws = wb.active
ws.title = "Data"

# Constants for calculations
iss_altitude = 408_000
earth_radius = 6_378_137

# Variables initialization
time_increase = 6
stop = False
very_first_data = True
time = datetime(2023,5,15,17,12,11)
yaw = 3
time_change = 0
total_time_elapsed = 0
total_yaw_change = 0

# Possible angular velocity values
ang_vel_vals = [(-3+num)/1000 for num in range(9)]
ang_vel_vals.remove(0)
ang_vel_vals.remove(0.001)

i = 1

# Main loop
while not stop:
    if not very_first_data:
        # Generate a random number to determine the angular velocity
        randting = random.randint(2, 3)
        if randting != 2:
            ang_vel = 0.001
        else:
            ang_vel = ang_vel_vals[random.randint(0, len(ang_vel_vals)-1)]
            
        # Calculate the change in yaw based on the angular velocity and time change
        yaw_change = ang_vel * time_change
        total_yaw_change += abs(yaw_change)
        
        # Check if the total yaw change has reached 360 degrees
        if total_yaw_change >= 360:
            stop = True
        
        # Update the yaw value and keep it within the range of 0 to 359 degrees
        yaw += yaw_change
        yaw %= 360
        
        # Calculate the linear velocity based on the angular velocity and altitude
        lin_vel = round((ang_vel * (iss_altitude + earth_radius)), 3)
        
        i += 1
        i %= 4
    else:
        # Initial values for the first data point
        ang_vel = -0.001
        lin_vel = -6786.14
        very_first_data = False
    
    # Append the data to the worksheet
    ws.append([time, yaw, ang_vel, lin_vel])
    
    # Increase time
    rand_num = random.randint(0,4)
    
    # Determine the new time increase based on the random number
    if rand_num == 0 or rand_num == 1:
        new_time_increase = time_increase
    elif rand_num == 2:
        new_time_increase = time_increase + 1
    else:
        new_time_increase = time_increase - 1
    
    # Update the time change and total time elapsed
    time_change = new_time_increase
    total_time_elapsed += time_change
    
    # Check if the total time elapsed has reached 10800 seconds (3 hours)
    if total_time_elapsed >= 10800:
        stop = True
        
    # Update the time by adding the new time increase
    time = time + timedelta(0, new_time_increase)

# Save the workbook to a file
wb.save("new_ting2.xlsx")
