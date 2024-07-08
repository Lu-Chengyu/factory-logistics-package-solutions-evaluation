# TLP Model: PC&L Costing Calculating & Solution Generating

## Overview

This project aims to optimize and ensure the feasibility of solutions for transportation logistics. It uses a combination of linear programming and simulation algorithms to generate solutions for loading trucks, pallets, and boxes efficiently.

## Project Structure

### Input Data

The input data includes:
- Daily volume of each part
- Frequency of delivery
- Distance from Location A to B
- Dimensions (length, width, height) of each part
- Available Box/Pallet/Truck this time
- Adhoc Box/Pallet/Truck if any

### Maintenance Data

Standard dimensions and investment costs for:
- Boxes
- Pallets
- Trucks

### Methodology

1. **For Simple Problems**: Generate solutions with Exhaustive Search.
2. **For Complicated Problems**: Generate solutions with Linear Programming and verify them with Simulation Algorithms if needed.

### Process

1. **Pick the Best-fit Box**: Using Combination Exhaustion on Box Type and Orientation (Excel).
2. **Pick the Best-fit Pallet**: 
   - Box to Pallet-Layer (100% occupied is a must).
     - Filled with a unique type of box.
     - Filled with miscellaneous types of boxes (same height).
   - Pallet-Layer to Pallet-Pile (Least Occupied Area) using Linear Programming and Simulation Algorithm (Python).
3. **Pick the Suitable Truck**: Cheapest possible combination using Linear Programming and Simulation Algorithm (Python).
4. **Calculate and Output Cost & Solution**.

### Output

The output includes:
- Type, number, and price of Truck/Pallet/Box.
- PC&L cost amortized to each part.
- Capacity utilization.
- Solution details & diagrams.

## Execution

To execute the model:
1. Ensure all input data is provided in the respective Excel sheets.
2. Run the Python script to generate the optimized solutions.
3. Review the output data for the best-fit boxes, pallets, and trucks along with the associated costs and capacity utilization.

## Dependencies

- Python 3.x
- numpy
- pulp
- itertools
- xlrd
- xlwt
- xlutils

## Instructions

1. Open and fill in the required data in the Excel templates provided.
2. Run the Python script to process the data and generate solutions.
3. The results will be saved in the Excel output file, which includes detailed solution diagrams and cost calculations.

## Troubleshooting

If you encounter issues, ensure that all Excel files are closed before running the script, as the script will try to access and modify these files.
