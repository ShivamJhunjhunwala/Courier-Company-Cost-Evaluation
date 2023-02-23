import pandas as pd
import math

# Reading EXCEL into DF
df_order_report = pd.read_excel("Company X - Order Report.xlsx")
df_sku_master = pd.read_excel("Company X - SKU Master.xlsx")
df_invoice = pd.read_excel("Courier Company - Invoice.xlsx")
df_pincodes = pd.read_excel("Company X - Pincode Zones.xlsx")
df_rates = pd.read_excel("Courier Company - Rates.xlsx").T

# Merge two df based on SKU value
df_order_sku = pd.merge(df_order_report, df_sku_master, how="inner", on="SKU")
df_order_sku.rename(columns={'ExternOrderNo': 'Order ID'}, inplace=True)

#
df_rates['type'] = df_rates.index
df_rates.reset_index(drop=True, inplace=True)
df_rates.rename(columns={0: 'Price'}, inplace=True)
df_rates[['type of shipment', 'zone', 'charge type']
         ] = df_rates.type.str.split("_", expand=True)
df_rates.drop(['type'], axis=1, inplace=True)
df_rates.drop_duplicates(inplace=True)
df_rates.sort_values(by=['type of shipment', 'zone'],
                     inplace=True, ignore_index=True)


output_df = pd.DataFrame()
output_df['Order ID'] = df_invoice['Order ID']
output_df['AWB Number'] = df_invoice['AWB Code']

xx = pd.merge(output_df, df_order_sku, on='Order ID')
xx['total weight'] = [value *
                      xx['Weight (g)'][index]/1000 for index, value in enumerate(xx['Order Qty'])]
weight_by_order_id = xx.groupby('Order ID').agg(weight=('total weight', 'sum'))

output_df = pd.merge(output_df, weight_by_order_id, on='Order ID')
output_df = pd.merge(output_df, df_invoice, on='Order ID')

output_df.drop(['AWB Code'], axis=1, inplace=True)
output_df['Weight slab as per X (KG)'] = [math.ceil(
    i/(0.5))*(0.5) for i in output_df['weight']]
output_df.rename(
    columns={"weight": 'Total weight as per X (KG)'}, inplace=True)
output_df['Weight slab charged by Courier Company (KG)'] = [math.ceil(
    i/(0.5))*(0.5) for i in output_df['Charged Weight']]
output_df.rename(columns={
                 'Charged Weight': 'Total weight as per Courier Company (KG)'}, inplace=True)
output_df['Delivery Zone as per X'] = df_pincodes['Zone']
output_df['Delivery Zone charged by Courier Company'] = df_invoice['Zone']
output_df = pd.merge(
    output_df, df_invoice[['Type of Shipment', 'Order ID']], on='Order ID')
lists = []
for index, value in enumerate(output_df['Type of Shipment_y']):
    count = math.ceil(
        output_df['Total weight as per Courier Company (KG)'][index]/(0.5))
    if 'Forward charges' == value:
        if output_df['Delivery Zone as per X'][index] == 'a':
            lists.append(
                round(df_rates['Price'][0] + (count-1)*df_rates['Price'][1], 1))
        elif output_df['Delivery Zone as per X'][index] == 'b':
            lists.append(
                round(df_rates['Price'][2] + (count-1)*df_rates['Price'][3], 1))
        elif output_df['Delivery Zone as per X'][index] == 'c':
            lists.append(
                round(df_rates['Price'][4] + (count-1)*df_rates['Price'][5], 1))
        elif output_df['Delivery Zone as per X'][index] == 'd':
            lists.append(
                round(df_rates['Price'][6] + (count-1)*df_rates['Price'][7], 1))
        elif output_df['Delivery Zone as per X'][index] == 'e':
            lists.append(
                round(df_rates['Price'][8] + (count-1)*df_rates['Price'][9], 1))
    else:
        if output_df['Delivery Zone as per X'][index] == 'a':
            lists.append(round(df_rates['Price'][0] + (count-1) *
                         df_rates['Price'][1] + (count)*df_rates['Price'][10], 1))
        elif output_df['Delivery Zone as per X'][index] == 'b':
            lists.append(round(df_rates['Price'][2] + (count-1) *
                         df_rates['Price'][3] + (count)*df_rates['Price'][12], 1))
        elif output_df['Delivery Zone as per X'][index] == 'c':
            lists.append(round(df_rates['Price'][4] + (count-1) *
                         df_rates['Price'][5] + (count)*df_rates['Price'][14], 1))
        elif output_df['Delivery Zone as per X'][index] == 'd':
            lists.append(round(df_rates['Price'][6] + (count-1) *
                         df_rates['Price'][7] + (count)*df_rates['Price'][16], 1))
        elif output_df['Delivery Zone as per X'][index] == 'e':
            lists.append(round(df_rates['Price'][8] + (count-1) *
                         df_rates['Price'][9] + (count)*df_rates['Price'][18], 1))
output_df['Expected Charge as per X (Rs.)'] = lists
output_df = pd.merge(
    output_df, df_invoice[['Billing Amount (Rs.)', 'Order ID']], on='Order ID')
output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] = [
    output_df['Expected Charge as per X (Rs.)'][i] - output_df['Billing Amount (Rs.)_y'][i] for i in range(len(output_df))]
output_df.drop(['Warehouse Pincode', 'Customer Pincode', 'Zone', 'Type of Shipment_y',
               'Type of Shipment_x', 'Billing Amount (Rs.)_x'], axis=1, inplace=True)
output_df.rename(columns={
                 'Billing Amount (Rs.)_y': 'Charges Billed by Courier Company (Rs.)'}, inplace=True)
output_df.loc[:, ['Order ID', 'AWB Number', 'Total weight as per X (KG)', 'Weight slab as per X (KG)', 'Total weight as per Courier Company (KG)',
                  'Weight slab charged by Courier Company (KG)',
                  'Delivery Zone as per X',
                  'Delivery Zone charged by Courier Company',
                  'Expected Charge as per X (Rs.)',
                  'Charges Billed by Courier Company (Rs.)',
                  'Difference Between Expected Charges and Billed Charges (Rs.)']]

correct = output_df[output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0][
    'Charges Billed by Courier Company (Rs.)'].sum()
less_then = output_df[output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0][
    'Difference Between Expected Charges and Billed Charges (Rs.)'].sum()
greater_then = output_df[output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0][
    'Difference Between Expected Charges and Billed Charges (Rs.)'].sum()
# output_df.to_excel("Output.xlsx")
data = [['Total orders where X has been correctly charged', output_df[output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0][
        'Difference Between Expected Charges and Billed Charges (Rs.)'].count(), correct],
        ['Total Orders where X has been overcharged',
        output_df[output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0][
        'Difference Between Expected Charges and Billed Charges (Rs.)'].count(), less_then],
        ['Total Orders where X has been undercharged',
        output_df[
        output_df['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0]
        ['Difference Between Expected Charges and Billed Charges (Rs.)'].count(), greater_then]]
new_df = pd.DataFrame(data, columns=['', 'Count', 'Amount (Rs.)'])
with pd.ExcelWriter("OUTPUT.xlsx") as writer:
    output_df.to_excel(writer, sheet_name="Calculation", index=False)
    new_df.to_excel(writer, sheet_name="Summary", index=False)
