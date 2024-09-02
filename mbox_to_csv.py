from datetime import datetime
import re
import mailbox
import pandas as pd


class MboxToCSV:

    def __init__(self, ipname):
        self.ipname = ipname

    def read(self):
        result = []
        mbox = mailbox.mbox(self.ipname)
        for message in mbox:
            result.append(
                {
                    "X-Gmail-Labels": message.get("X-Gmail-Labels"),
                    "To": message.get("To"),
                    "From": message.get("From"),
                    "Subject": message.get("Subject"),
                    "Date": message.get("Date"),
                }
            )
        return result

    def execution_timer(func):

        def timer(self, *args, **kwargs):
            start = datetime.now()
            print(f"""Start Time: {start.strftime("%Y-%m-%d %H:%m:%S")}""")
            func(self, *args, **kwargs)
            end = datetime.now()
            print(f"""End Time: {start.strftime("%Y-%m-%d %H:%m:%S")}""")
            print(f"""Total Time: {(end-start).seconds} seconds""")

        return timer

    def transform_date(self, df):
        # pd.to_datetime(df["Date"], format='mixed', cache=True, errors="coerce")
        # Fix for multiple format
        # Proposed convert and format to date which are good data, merge and get bad data
        # pd.to_datetime(df["Date"], format='mixed', cache=True, errors="coerce").dt.strftime("%y-%m-%d")
        pass

    @execution_timer
    def generate_file(self, opname):
        df = pd.DataFrame(self.read())
        df["From"] = df["From"].apply(
            lambda x: re.findall(r"<(.*?)>", x) if "<" in x else x
        )
        df.dropna(subset=["From"], inplace=True)
        pd.to_datetime(
            df["Date"], format="mixed", cache=True, errors="coerce", utc=True
        ).dt.strftime("%y-%m-%d")
        # self.transform_date(df) # In Future implementation for more roboust approach
        # TODO: create sheet for inbox and send and other instead 1
        df.to_csv(opname, index=False)


if __name__ == "__main__":
    ipname = "/Users/dmeow/Downloads/All mail Including Spam and Trash-001.mbox"
    opname = "report.csv"
    obj = MboxToCSV(ipname)
    obj.generate_file(opname)
    print("complete.........")
