import os
import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl import Workbook

class ExpenseBuilder:
    def __init__(self, folderPath):
        self.folderPath = folderPath
        self.fileList = []
        self.parsedData = []

    def read_folder(self):
        # List all files in the folder - leaving for testing & development purposes 
        files = os.listdir(self.folderPath)
        print(f"Files in '{self.folderPath}':")
        for file in files:
            print(file)
            self.fileList.append(file)
        
    def getActivityData(self):
        #Parse through files, please ensure your provider is included in file name e.g. amex - amex_activity.xlsx
        for file in self.fileList:
            file_path = os.path.join(self.folderPath, file)
            if file.endswith('.csv'):
                #Capital One
                if "c1" in file: 
                    df = self.parseCapitalOneCSV(file_path)
                    self.parsedData.append(df)
            elif file.endswith('.xlsx'):
                #American Express Card 
                if "amex" in file: 
                    df = self.parseAmericanExpressExcel(file_path)
                    self.parsedData.append(df)
                else: 
                    print("using different method")

    def buildExpenseSheet(self):
        """
        This will build the expense sheet with data in self.parsedData 
        
        Columns: 
        Transaction Date - Formatted: %Y-%m-%d
        Description 
        Category 
        Amount 

        Returns:
            None 
        """
        if self.parsedData:
            # Concat SpreadSheet Data  
            df = pd.concat(self.parsedData, ignore_index=True)
            
            # Sort by Transaction Date 
            df = df.sort_values(by='Transaction Date')

            # Calculate Sum
            total_sum = df['Amount'].sum()

            # Append a "Total" row
            total_row = pd.DataFrame({
                'Total': [total_sum]  
            })
            df = pd.concat([df, total_row], ignore_index=True)

            # Add "$" sign to amount & total column
            df['Amount'] = df['Amount'].apply(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")
            df['Total'] = df['Total'].apply(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")

            #Output Path 
            output_path = os.path.join(self.folderPath, "ExpenseSheet.xlsx")

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write the DataFrame to Excel
                df.to_excel(writer, index=False, sheet_name="Expenses")
                worksheet = writer.sheets["Expenses"]
                
                # Center align text and adjust column widths
                for col in worksheet.columns:
                    max_length = 0
                    col_letter = col[0].column_letter  # Get column letter (e.g., A, B, C)
                    for cell in col:
                        # Center text
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        # Calculate max content length
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = max_length + 2  # Add some padding to the width
                    worksheet.column_dimensions[col_letter].width = adjusted_width

    def parseCapitalOneCSV(self, file_path):  
        print(f"\nReading Capital One CSV file: {file_path}")
      
        try: 
            df = pd.read_csv(file_path)

            #Drop rows where 'Debit' is nil - We do not want to track payments made to Capital One.  
            df = df.dropna(subset=['Debit']) 

            df.rename(columns={
                'Debit': 'Amount',
            }, inplace=True)

            #In 2.0, make this dynamic
            df['Payment Method'] = "Capital One"

            return df[['Transaction Date', 'Description', 'Category', 'Payment Method', 'Amount']]

        except Exception as e:
            print(f"{e}")


    def parseAmericanExpressExcel(self, file_path):
        '''
        American Express .xlsx are formatted weirdly. 

        '''
        print(f"\nReading American Express Excel file: {file_path}")
        try: 
            #Skip 6 rows as we do not need that information when parsing. 
            df = pd.read_excel(file_path, skiprows=6)

            #Drop rows where 'Category' is nil - We do not want to track payments made to American Express. 
            df = df.dropna(subset=['Category']) 
        
            df.rename(columns={
                'Date': 'Transaction Date',

            }, inplace=True)

            #Format Date 
            df['Transaction Date'] = pd.to_datetime(df['Transaction Date'])
            df['Transaction Date'] = df['Transaction Date'].dt.strftime('%Y-%m-%d')

            df['Payment Method'] = "American Express"

            return df[['Transaction Date', 'Description', 'Category', 'Payment Method', 'Amount']]
            
        except Exception as e:
            print(f"{e}")

    def parseRobinhoodGold(self): 
        '''
        Implement OCR 
        '''
        pass




if __name__ == "__main__":
    #path of folder with .xlsx & .cs
    folderPath = 'data_folder'  

    expense_builder = ExpenseBuilder(folderPath)
    expense_builder.read_folder()
    expense_builder.getActivityData()
    expense_builder.buildExpenseSheet()