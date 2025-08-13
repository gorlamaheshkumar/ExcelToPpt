import time
import os

print("--- Python Test Script Starting Up ---")
print(f"This process will now run forever. Process ID: {os.getpid()}")
print("You should see a 'Hello world' message in the logs every 10 seconds.")
print("---------------------------------------------------")

# Infinite loop to keep the container running
while True:
    print("Hello world! The script is alive.")
    # The script will print this message every 10 seconds
    time.sleep(10)