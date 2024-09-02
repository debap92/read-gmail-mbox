from datetime import datetime
import re
import mailbox
import pandas as pd


class MboxToExcel:

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
        data = self.read()
        df = pd.DataFrame(data)
        df.dropna(subset=["From"], inplace=True)
        df["From"] = (
            df["From"]
            .astype(str)
            .apply(lambda x: re.findall(r"<(.*?)>", x)[0] if "<" in x else x)
        )

        df["Date"] = df["Date"].apply(
            lambda x: re.findall(r"\s*\d+\s+\w+\s+\d{4}", x)[0].strip().lstrip("0")
        )
        # self.transform_date(df) # In Future implementation for more robust approach
        # TODO: create sheet for inbox and send and other instead 1
        with pd.ExcelWriter(opname) as writer:
            df.groupby("From")["From"].count().reset_index(name="count").sort_values(
                ["count"], ascending=False
            ).to_excel(writer, sheet_name="Count", index=False)
            df.to_excel(writer, sheet_name="Details", index=False)


if __name__ == "__main__":
    ipname = "/Users/dmeow/Downloads/All mail Including Spam and Trash-001.mbox"
    opname = "report.xlsx"
    obj = MboxToExcel(ipname)
    obj.generate_file(opname)
    print("complete.........")
