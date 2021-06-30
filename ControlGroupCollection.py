import pandas as pd
import time
import sys


class DuplicateRecordError(Exception):
    """Exception raised duplicate record found in parent collection

    Attributes:
        salary -- input salary which caused the error
        message -- explanation of the error
    """
    def __init__(self, message="Found a matching record in the parent control_group"):
        self.message = message
        super().__init__(self.message)
# --------------------------  End of Exception Definitions --------------------------------


class ControlGroupMasterDataSet:
    def __init__(self):
        self.project_collection = pd.DataFrame()

    def get_collections(self):
        return self.project_collection

    def commit_control_group(self, control_group):
        try:
            self.project_collection.append(control_group, ignore_index=True)
            return True
        except Exception:
            print("Unable to add control group to parent control_group")
            e = sys.exc_info()[0]
            print("<p>Error: %s</p>" % e)
            exit(2)
            return False

# --------------------------------------  End of Class ControlGroupDataSet  -----------------------------------------