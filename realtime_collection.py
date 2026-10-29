import psutil
import openpyxl
import time

# Function to collect system data
def collect_system_data():
    cpu_usage = psutil.cpu_percent(interval=1)
    ram_usage = psutil.virtual_memory().percent
    battery = psutil.sensors_battery()
    power_plugged = battery.power_plugged if battery else False
    network_speed = psutil.net_io_counters()
    return cpu_usage, ram_usage, 1 if power_plugged else 0, network_speed


# Create an Excel workbook and add a worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.append(["Metric", "Average Value"])

try:
    cpu_sum = ram_sum = power_sum = upload_sum = download_sum = 0
    count = 0

    print("Calculating average values... ", end="")
    for _ in range(10):  # 15 minutes (900 seconds)
        cpu, ram, power, network = collect_system_data()

        cpu_sum += cpu
        ram_sum += ram
        power_sum += power
        upload_sum += network.bytes_sent
        download_sum += network.bytes_recv

        count += 1
        time.sleep(1)  # Wait for 1 second
        print(".", end="", flush=True)  # Display loading dots
    
    # Calculate average values
    avg_cpu = cpu_sum / count
    avg_ram = ram_sum / count
    avg_power = power_sum / count
    avg_upload_speed = (upload_sum / count) / (1024 * 1024)  # Convert to MB
    avg_download_speed = (download_sum / count) / (1024 * 1024)  # Convert to MB

    # Store average values in the Excel file
    worksheet.append(["CPU Usage (%)", avg_cpu])
    worksheet.append(["RAM Usage (%)", avg_ram])
    worksheet.append(["Power Plugged", avg_power])
    worksheet.append(["Upload Speed (MB)", avg_upload_speed])
    worksheet.append(["Download Speed (MB)", avg_download_speed])

    # Save the Excel file
    workbook.save("system_metrics.xlsx")

    print("\nData stored into the Excel file 'system_metrics.xlsx'.")

except KeyboardInterrupt:
    print("\nData collection stopped by user.")

