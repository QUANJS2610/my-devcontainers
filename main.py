"""
This is the main file for the project. This file will be used to run the project.
The main() function is the entry point for the project.
"""

# Importing the required libraries for the project
import sys

# Adding the code directory to the path so that the code can be imported
sys.path.append('./code')
from Extract_Data_from_Gmail import get_data_from_Gmail

# Importing the required libraries for the project
def main():
    get_data_from_Gmail()

# The following is the standard boilerplate that calls the main() function. 
if __name__ == "__main__":
    main()