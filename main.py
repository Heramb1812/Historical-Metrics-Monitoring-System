import openpyxl
from datetime import datetime, timedelta
import random
from openpyxl import load_workbook


# Provided 15-minute averages
avg_cpu_15min_initial = 1.43
avg_ram_15min_initial = 71.22
avg_power_plugged_15min_initial = 0.75
avg_upload_speed_15min_initial = 78.88
avg_download_speed_15min_initial = 2082.53
avg_network_latency_15min_initial = 10  # Simulated initial network latency in ms
avg_disk_usage_15min_initial = 40.0     # Simulated initial disk usage in %
avg_memory_usage_15min_initial = 60.0   # Simulated initial memory usage in %
avg_temperature_15min_initial = 40.0    # Simulated initial temperature in °C
avg_gpu_usage_15min_initial = 40.0      # Simulated initial GPU usage in %
avg_battery_remaining_15min_initial = 75.0  # Simulated initial battery remaining in %
avg_fan_speed_15min_initial = 2000.0    # Simulated initial fan speed in RPM
avg_cpu_clock_speed_15min_initial = 2.5  # Simulated initial CPU clock speed in GHz
avg_disk_temperature_15min_initial = 35.0  # Simulated initial disk temperature in °C
avg_disk_read_speed_15min_initial = 50.0  # Simulated initial disk read speed in MB/s
avg_disk_write_speed_15min_initial = 30.0  # Simulated initial disk write speed in MB/s
avg_cpu_temperature_15min_initial = 50.0  # Simulated initial CPU temperature in °C
avg_gpu_temperature_15min_initial = 60.0  # Simulated initial GPU temperature in °C
avg_frame_rate_15min_initial = 144.0  # Simulated initial frame rate

# Get today's date and set starting date to 90 days ago
today = datetime.today().date()
start_date = today - timedelta(days=90)

# +1 from the current date and going back to the past for 90 days which is 91 days 90+1 = 91

# Create an Excel workbook and add a worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.append(["Time", "CPU Usage (%)", "RAM Usage (%)", "Power Plugged", "Upload Speed (MB)", "Download Speed (MB)", "Network Latency (ms)", "Disk Usage (%)", "Memory Usage (%)", "Temperature (°C)", "GPU Usage (%)", "Battery Remaining (%)", "Fan Speed (RPM)", "CPU Clock Speed (GHz)", "Disk Temperature (°C)", "Disk Read Speed (MB/s)", "Disk Write Speed (MB/s)", "CPU Temperature (°C)", "GPU Temperature (°C)", "Refresh Rate (Hz)"])


try:
    # Simulate data for each day
    while start_date <= today:
        # Calculate rate of change per day (with variation)
#virations should be in the identical  manners
        cpu_rate_of_change = random.uniform(-0.5, 0.5)
        ram_rate_of_change = random.uniform(-2, 2)
        power_plugged_rate_of_change = random.uniform(-0.05, 0.05)
        upload_rate_of_change = random.uniform(-5, 5)
        download_rate_of_change = random.uniform(-50, 50)
        network_latency_rate_of_change = random.uniform(-2, 2)  # Variation for network latency
        disk_usage_rate_of_change = random.uniform(-1, 1)       # Variation for disk usage
        memory_usage_rate_of_change = random.uniform(-1, 1)     # Variation for memory usage
        temperature_rate_of_change = random.uniform(-1, 1)      # Variation for temperature
        gpu_usage_rate_of_change = random.uniform(-1, 1)        # Variation for GPU usage
        battery_remaining_rate_of_change = random.uniform(-1, 1)  # Variation for battery remaining
        fan_speed_rate_of_change = random.uniform(-50, 50)      # Variation for fan speed
        cpu_clock_speed_rate_of_change = random.uniform(-0.1, 0.1)  # Variation for CPU clock speed
        disk_temperature_rate_of_change = random.uniform(-1, 1)  # Variation for disk temperature
        disk_read_speed_rate_of_change = random.uniform(-5, 5)   # Variation for disk read speed
        disk_write_speed_rate_of_change = random.uniform(-3, 3)  # Variation for disk write speed
        cpu_temperature_rate_of_change = random.uniform(-2, 2)   # Variation for CPU temperature
        gpu_temperature_rate_of_change = random.uniform(-2, 2)   # Variation for GPU temperature


        avg_frame_rate = avg_frame_rate_15min_initial  # Since refresh rate is constant at 144Hz
         #algorithm works for the whole project 
        avg_cpu = avg_cpu_15min_initial + cpu_rate_of_change
        avg_ram = avg_ram_15min_initial + ram_rate_of_change
        avg_power_plugged = avg_power_plugged_15min_initial + power_plugged_rate_of_change
        avg_upload_speed = avg_upload_speed_15min_initial + upload_rate_of_change
        avg_download_speed = avg_download_speed_15min_initial + download_rate_of_change
        avg_network_latency = avg_network_latency_15min_initial + network_latency_rate_of_change
        avg_disk_usage = avg_disk_usage_15min_initial + disk_usage_rate_of_change
        avg_memory_usage = avg_memory_usage_15min_initial + memory_usage_rate_of_change
        avg_temperature = avg_temperature_15min_initial + temperature_rate_of_change
        avg_gpu_usage = avg_gpu_usage_15min_initial + gpu_usage_rate_of_change
        avg_battery_remaining = avg_battery_remaining_15min_initial + battery_remaining_rate_of_change
        avg_fan_speed = avg_fan_speed_15min_initial + fan_speed_rate_of_change
        avg_cpu_clock_speed = avg_cpu_clock_speed_15min_initial + cpu_clock_speed_rate_of_change
        avg_disk_temperature = avg_disk_temperature_15min_initial + disk_temperature_rate_of_change
        avg_disk_read_speed = avg_disk_read_speed_15min_initial + disk_read_speed_rate_of_change
        avg_disk_write_speed = avg_disk_write_speed_15min_initial + disk_write_speed_rate_of_change
        avg_cpu_temperature = avg_cpu_temperature_15min_initial + cpu_temperature_rate_of_change
        avg_gpu_temperature = avg_gpu_temperature_15min_initial + gpu_temperature_rate_of_change


        # Calculate the refresh rate
        avg_refresh_rate = 144  # Refresh rate remains constant at 144 Hz
        
        # Append simulated averages to the Excel file
        timestamp = start_date.strftime("%Y-%m-%d %I:%M:%S %p")
        worksheet.append([timestamp, avg_cpu, avg_ram, avg_power_plugged, avg_upload_speed, avg_download_speed,
                          avg_network_latency, avg_disk_usage, avg_memory_usage, avg_temperature, avg_gpu_usage,
                          avg_battery_remaining, avg_fan_speed, avg_cpu_clock_speed, avg_disk_temperature,
                          avg_disk_read_speed, avg_disk_write_speed, avg_cpu_temperature, avg_gpu_temperature,
                          f"{avg_refresh_rate} Hz"])  # Append refresh rate with " Hz"

        # Increment the date by one day
        start_date += timedelta(days=1)


# till we have finished the explantion.................................
        







    # Apply percentage formatting to relevant columns as for the worksheet appending
    percentage_columns = [4, 5, 8, 9, 10, 12, 16, 17, 18]
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=4, max_col=21):  # Changed max_col to 21
        for idx, cell in enumerate(row):
            if idx in percentage_columns:
                cell.number_format = '0.00%'

    # Format the "Refresh Rate (Hz)" column to not include a comma and percentage symbol bedefault we need to set it 
    for cell in worksheet["U"][1:]:
        cell.number_format = '0.00'

    # Save the Excel file as per your require name 
    workbook.save("captured.xlsx")

    print("Simulated Power BI data for 90 days stored into the Excel file 'captured.xlsx'.")


except KeyboardInterrupt:
    print("\nData simulation stopped by user.")
    


