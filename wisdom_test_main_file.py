from create_dataset import create_dataset
import warnings

if __name__ == '__main__':
    file_loc = 'Dummy Data for final assignment.xlsx'
    print("--- Starting ---")
    warnings.filterwarnings("ignore")
    create_dataset(file_loc)
    print("--- Ending ---")


# -------------------------------------------------------
# -------------------------------------------------------

