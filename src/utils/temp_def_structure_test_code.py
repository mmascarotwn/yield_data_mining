# test structure test code

# This '.py' file serves the sole purpose of side dvelopment
# The used file must go in 'examples/mobility/gmsh_mos2d_characterisation.py'

import os
import csv
import numpy as np

from scipy.stats import linregress
import matplotlib.pyplot as plt

from tkinter import Tk
from tkinter.filedialog import askopenfilenames, asksaveasfilename

import devsim.python_packages.simple_physics as simple_physics


# Generic description: 
# Rtot at 4 different channel lengths, Rc extracted as the intercept on the y-axis
# Take the data from 4 different external .csv file with Rtot at different channel lenghts
# Merge them together & add the channel length
# Use small Vds (0.1V)


# Function to merge CSV files
# TO-DO: give automatically the channel length (to be added at priori during the simulation)
def merge_csv_files_gui():
    """
    Allows the user to select 4 different .csv files via a GUI and merges their content into a single .csv file.
    Adds a new column "Lch" to the output file.
    """
    # Step 1: Open a file selection dialog to select 4 .csv files
    Tk().withdraw()  # Hide the root Tkinter window
    print("Please select 4 .csv files to merge.")
    file_paths = askopenfilenames(
        title="Select 4 CSV Files",
        filetypes=[("CSV Files", "*.csv")],
        multiple=True
    )

    # Ensure exactly 3 files are selected
    if len(file_paths) != 4:
        print("Error: You must select exactly 3 .csv files.")
        return

    # Step 2: Open a save dialog to specify the output file
    output_file = asksaveasfilename(
        title="Save Merged CSV File As",
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv")]
    )

    if not output_file:
        print("Error: No output file specified.")
        return

    # Step 3: Merge the content of the selected files into the output file
    try:
        with open(output_file, mode="w", encoding="utf-8", newline="") as output_csv:
            writer = None  # CSV writer object
            
            for idx, file_path in enumerate(file_paths):

                with open(file_path, mode="r", encoding="utf-8") as input_csv:
                    reader = csv.reader(input_csv)
                    headers = next(reader)  # Read the header row

                    if writer is None:
                        # Write the header row to the output file and add "Lch"
                        writer = csv.writer(output_csv)
                        # headers.append("Lch")  # Add the "Lch" column to the headers
                        writer.writerow(headers)

                    for row in reader:
                        writer.writerow(row)  # Write each row to the output file
        print(f"Merged content saved to '{output_file}'.")
    except Exception as e:
        print(f"An error occurred while merging files: {e}")
    return output_file


# Function to extract 2*Rc 'contact resistance' from the TLM
# 1) Rc_tot = intercept on the y-axis at L=0
# Rc_tot = 2*Rc (is 2 times the contact resistance of a single pad, assumed to be the same for both pads)
def rc_tlm(file_path = None):
    """
    Reads a .csv file, extracts 'Rtot' and 'Lch' values, and calculates the intercept (Rc) on the y-axis for Lch=0.

    Args:
        file_path (str): Path to the .csv file.

    Returns:
        Rc (float): The intercept on the y-axis (contact resistance).
    """
    try:

        # Step 0: Load files
        # file_path = merge_csv_files_gui()

        # Step 1: Read the .csv file and extract 'Rtot' and 'Lch' values
        lch_values = []
        # lch_values_cm = []
        rtot_values = []
        rows = []

        with open(file_path, mode='r', encoding='utf-8') as csv_file:
            reader = csv.DictReader(csv_file)
            headers = reader.fieldnames

            # Ensure the 'Rc' column exists in the headers
            if 'Rc' not in headers:
                headers.append('Rc')

            for row in reader:
                try:
                    lch_values.append(float(row['Lch']))
                    # lch_values_cm.append(float(row['Lch']) * (1e-7))
                    rtot_values.append(float(row['Rtot']))
                    rows.append(row)
                except KeyError:
                    print("Error: Required columns ('Lch', 'Rtot') not found in the CSV file.")
                    return None
                except ValueError:
                    print("Error: Invalid data in 'Lch' or 'Rtot' columns.")
                    return None
                
        # Step 2: Prompt the user for a scaling factor for the y-axis
        try:
            scaling_factor = float(input("Enter the scaling factor for the y-axis (Ids): "))
        except ValueError:
            print("Invalid scaling factor. Using default value of 1.")
            scaling_factor = 1.0        

        # Convert lists to NumPy arrays
        lch_values = np.array(lch_values)
        # lch_values_cm = np.array(lch_values_cm)
        rtot_values = np.array(rtot_values) * scaling_factor  # Apply the scaling factor to Rtot

        # Step 2: Perform linear regression to find Rc (intercept at Lch=0)
        slope, intercept, r_value, p_value, std_err = linregress(lch_values, rtot_values)
        Rc_tot = intercept  # Intercept is 2*Rc
        Rc = Rc_tot / 2  # Divide by 2 to get Rc (assuming same contact gemetry and physical behaviour)

        # Step 3: Write Rc_tot to all rows under the 'Rc' column
        for row in rows:
            row['Rc'] = Rc_tot

        # Step 4: Save the updated data back to the .csv file
        with open(file_path, mode='w', encoding='utf-8', newline='') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=headers)
            writer.writeheader()  # Write the header row
            writer.writerows(rows)  # Write all rows
        print(f"Updated .csv file saved to '{file_path}' with Rc_tot = {Rc_tot:.2f}.")

        # Step 5: Plot the data and the regression line
        plt.figure(figsize=(8, 6))
        plt.scatter(lch_values, rtot_values, label="Data Points", color="blue")
        # TO-DO: make this automatic every time scaling factor is = 1
        plt.plot(lch_values, slope * lch_values + intercept, label=f"Fit: Rtot = {slope:.2f}*Lch + {intercept:.2f} (Ω*cm)", color="red")#  (no scaling factor)
        # plt.plot(lch_values, slope * lch_values + intercept, label=f"Fit: Rtot = {slope:.2f}*Lch + {intercept:.2f} (kΩ*um)", color="red") # (scaled factor by 1e1 = 10)
        plt.axhline(y=Rc_tot, color="green", linestyle="--", label=f"Rc (Intercept) = {Rc_tot:.2f} (Ω*cm)") # (no scaling factor)
        # plt.axhline(y=Rc_tot, color="green", linestyle="--", label=f"Rc (Intercept) = {Rc_tot:.2f} (kΩ*um)") # scaled factor by 1e1 = 10
        plt.xlim(left=0)
        plt.xlabel("Lch (nm)")
        # TO-DO: make this automatic every time scaling factor is = 1
        plt.ylabel("Rtot (Ω*cm)") # no scaling factor
        # plt.ylabel("Rtot (kΩ*um)") # scaled factor by 1e1 = 10
        plt.title("Linear Fit to Extract Rc")
        plt.legend()
        plt.grid(False)
        plt.show()

        return Rc_tot, Rc

    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return Rc_tot, Rc


# Function to extract 'Rch' the 'channel resistance' from the TLM
# 2) Rch = Rtot - Rc_tot; 
def rch_tlm(input_file, rc_tot):
    """
    Reads a .csv file, calculates 'Rch' for each row by subtracting 'Rc_tot' from 'Rtot',
    and saves the updated data to a new .csv file.

    Args:
        input_file (str): Path to the input .csv file.
        output_file (str): Path to the output .csv file.
        rc_tot (float): The Rc_tot value to subtract from Rtot.
    """
    try:
        # Step 1: Read the input .csv file
        rows = []
        rch_values = []  # List to store Rch values
        with open(input_file, mode='r', encoding='utf-8') as csv_file:
            reader = csv.DictReader(csv_file)
            headers = reader.fieldnames

            # Add 'Rch' to the headers if it doesn't already exist
            if 'Rch' not in headers:
                headers.append('Rch')

            for row in reader:
                try:
                    # Step 2: Calculate 'Rch' = 'Rtot' - 'Rc_tot'
                    rtot = float(row['Rtot'])
                    rch = rtot - rc_tot
                    row['Rch'] = rch  # Add the calculated 'Rch' to the row
                    rch_values.append(float(rch))  # Store the Rch value in the list
                except ValueError:
                    print(f"Warning: Invalid data in row: {row}")
                    row['Rch'] = None  # Set 'Rch' to None if calculation fails
                    rch_values.append(None)  # Append None to the list
                rows.append(row)

        # Step 3: Save the updated data to the output .csv file
        with open(input_file, mode='w', encoding='utf-8', newline='') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=headers)
            writer.writeheader()  # Write the header row
            writer.writerows(rows)  # Write all rows
        print(f"Updated .csv file saved to '{input_file}'.")

        return rch_values  # Return the list of Rch values

    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


# Function to extract 'Rsh' the 'sheet resistance' from the TLM
# 3) Rsh = Rch / Lch
def rsh_tlm(input_file):
    """
    Reads a .csv file, calculates 'Rch' for each row by subtracting 'Rc_tot' from 'Rtot',
    and saves the updated data to a new .csv file.

    Args:
        input_file (str): Path to the input .csv file.
        output_file (str): Path to the output .csv file.
        rc_tot (float): The Rc_tot value to subtract from Rtot.
    """
    try:
        # Step 1: Read the input .csv file
        lch_values = []
        rows = []
        rsh_values = []  # List to store Rch values
        with open(input_file, mode='r', encoding='utf-8') as csv_file:
            reader = csv.DictReader(csv_file)
            headers = reader.fieldnames

            # Add 'Rch' to the headers if it doesn't already exist
            if 'Rsh' not in headers:
                headers.append('Rsh')

            for row in reader:
                try:
        # Step 2: Calculate 'Rsh' = 'Rch'/'Lch'
                    rch = float(row['Rch'])
                    # lch_values.append(float(row['Lch']))
                    Lch = float(row['Lch'])
                    Lch_cm = Lch * 1e-7  # Convert Lch from nm to cm
                    rsh = rch/Lch_cm
                    row['Rsh'] = rsh  # Add the calculated 'Rch' to the row
                    rsh_values.append(float(rsh))  # Store the Rch value in the list
                except ValueError:
                    print(f"Warning: Invalid data in row: {row}")
                    row['Rsh'] = None  # Set 'Rch' to None if calculation fails
                    rsh_values.append(None)  # Append None to the list
                rows.append(row)

        # Step 3: Save the updated data to the output .csv file
        with open(input_file, mode='w', encoding='utf-8', newline='') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=headers)
            writer.writeheader()  # Write the header row
            writer.writerows(rows)  # Write all rows
        print(f"Updated .csv file saved to '{input_file}'.")

        # Step 4: Plot the data and the regression line
        # lch_values = np.array(lch_values)         # Convert lists to NumPy arrays
        # rsh_values = np.array(rsh_values) # lch_values_cm = np.array(lch_values_cm)
        # plt.figure(figsize=(8, 6))
        # plt.scatter(lch_values, rsh_values, label="Data Points", color="blue")
        # plt.plot(lch_values, slope * lch_values + intercept, label=f"Fit: Rtot = {slope:.2f}*Lch + {intercept:.2f} (Ω*cm)", color="red")
        # plt.axhline(y=Rc_tot, color="green", linestyle="--", label=f"Rc (Intercept) = {Rc_tot:.2f} (Ω*cm)")
        # plt.xlim(left=0)
        # plt.xlabel("Lch (nm)")
        # plt.ylabel("Rsh (Ω*sqr)")
        # plt.title("Rsh vs Lch")
        # plt.legend()
        # plt.grid(True)
        # plt.show()

        return rsh_values  # Return the list of Rch values

    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


# Function to extract the Channel Mobility from Id@Vg
# This function is a placeholder and needs to be implemented based on the specific mobility calculation method.
# ucon = 1/(q * ns * Rsh); ns = (Cin * Vov)/q (near the source; Rsh = sheet resistance extracted through TLM
# Rsh is extracted from the TLM
def mobility_con(file_path):
    dibl = None

    return dibl


# TO-DO: compute resistances and mobility from TLMs # MM
def compute(a=None):

    # Step 1: Select input files and generate merged file
    file_path = merge_csv_files_gui()

    # Step 2: Compute parameters & store
    Rc_tot,_ = rc_tlm(file_path)
    Rch = rch_tlm(file_path, Rc_tot)
    Rsh = rsh_tlm(file_path)

 

###
### Test the functions
###

compute()

# merge_csv_files_gui()

# Example usage
# file_path = '/Users/macbookpro/Desktop/TLM_output_tofill.csv'  # Replace with the path to your .csv file
# Rc_tot,_ = rc_tlm(file_path)
# Rch = rch_tlm(file_path, Rc_tot)
# Rsh = rsh_tlm(file_path)
# if Rc_tot is not None:
#     print(f"Calculated Rc_tot (Contact Resistance): {Rc_tot:.2f}")
# if Rch is not None:
#     print(f"Calculated Rch (Channel Resistance): {Rch}")
# if Rsh is not None:
#     print(f"Calculated Rsh (Sheet Resistance): {Rsh}")

#######
# file_path = '/Users/macbookpro/Desktop/id_vds.csv'  # Replace with the path to your CSV file

# n2D, r_tot = r_tot(file_path)
# test = compute()


# if ioff:
#     current_at_0 = i_off
#     print(f"Ioff at Vgs = 0V: {current_at_0}")


# if n2D:
#     print(f"n2D is: {n2D:.2e}", "cm^-2") # cm^-2
# if r_tot:
#     print(f"Rtot is: {r_tot:.2e}", "kΩ*um") # kOhm * um