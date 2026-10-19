import pandas as pd

def generate_supplier_scorecard():
    # 1. Sample Procurement & Delivery Dataset
    data = {
        'PO_Number': ['PO-1001', 'PO-1002', 'PO-1003', 'PO-1004', 'PO-1005', 'PO-1006'],
        'Supplier_Name': ['Apex Logistics', 'Apex Logistics', 'Beta Industrial', 'Beta Industrial', 'Delta Steel', 'Delta Steel'],
        'Promised_Days': [5, 5, 7, 7, 4, 4],
        'Actual_Days': [5, 6, 7, 9, 3, 4],
        'Order_Value': [15000, 22000, 8500, 12000, 31000, 18000]
    }
    
    df = pd.DataFrame(data)
    
    # 2. Calculate On-Time Status (Actual Days <= Promised Days)
    df['On_Time'] = df['Actual_Days'] <= df['Promised_Days']
    
    # 3. Aggregate Performance by Supplier
    scorecard = df.groupby('Supplier_Name').agg(
        Total_Orders=('PO_Number', 'count'),
        On_Time_Orders=('On_Time', 'sum'),
        Total_Spend=('Order_Value', 'sum'),
        Avg_Lead_Time=('Actual_Days', 'mean')
    ).reset_index()
    
    # 4. Calculate On-Time Delivery (OTD) Percentage
    scorecard['OTD_Percentage'] = (scorecard['On_Time_Orders'] / scorecard['Total_Orders']) * 100
    
    # 5. Assign Risk Tier based on OTD
    def assign_risk(otd):
        if otd >= 90:
            return 'Low Risk'
        elif otd >= 75:
            return 'Medium Risk'
        else:
            return 'High Risk'
            
    scorecard['Risk_Rating'] = scorecard['OTD_Percentage'].apply(assign_risk)
    
    # Display the final scorecard report
    print("==================================================")
    print("         SUPPLIER PERFORMANCE SCORECARD           ")
    print("==================================================")
    print(scorecard.to_string(index=False))
    print("==================================================")

if __name__ == "__main__":
    generate_supplier_scorecard()