import numpy as np
from Get_Data import get_data_for_calculation  # Import data retrieval function from Get_Data.py

# Function to calculate the result using the retrieved data
def calculate_result(data):
    """
    Perform the SUMPRODUCT-like calculations and return the result.
    """
    # Retrieve the data for specific columns
    conflict_averse_range = np.array([float(x) for x in data["conflict_averse"]])
    overall_performance_range = np.array([float(x) for x in data["overall_performance"]])
    flourish_range = np.array([float(x) for x in data["flourish"]])
    cognitive_value = float(data["cognitive"][0])  # Assuming single value for cognitive (can adjust for row-based)
    
    # Retrieve the data for the ranges (D6:AG6, AH6:AR6, AU6:BH6, BW6)
    d_ag_range = np.array([float(x) for x in data["d_ag"]])
    ah_ar_range = np.array([float(x) for x in data["ah_ar"]])
    au_bh_range = np.array([float(x) for x in data["au_bh"]])
    bw_value = float(data["bw"])
    
    # Placeholder values for weights (replace with dynamic lookups if needed)
    conflict_averse_weight = 0
    overall_performance_weight =0
    flourish_weight = 0
    cognitive_weight = 0

    # Perform SUMPRODUCT-like calculation (same logic as in the Excel formula)
    result = (np.dot(d_ag_range, conflict_averse_range)+
              np.dot(ah_ar_range, overall_performance_range) +
              np.dot(au_bh_range, flourish_range) +
              bw_value * cognitive_value) / 100

    # Apply additional logic for "Location" and "LJI Knowledge"
    location_value = float(data["location"][0])  # Assuming single value for Location (can adjust)
    lji_knowledge_value = float(data["lji_knowledge"][0])  # Assuming single value for LJI Knowledge (can adjust)
    
    result += location_value + lji_knowledge_value

    return result

# Main script to run the data retrieval and calculation logic
def main():
    # Step 1: Retrieve the data from the sheet using Get_Data.py functions
    data = get_data_for_calculation()

    # Step 2: Perform the calculations on the retrieved data
    result = calculate_result(data)

    # Print or write back the calculated result
    print(f"Calculated Result: {result}")

if __name__ == "__main__":
    main()
