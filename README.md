
# Courier-Company-Cost-Evaluation

This project helps us to validate bills given by a specific courier company for a package. This bill depends mainly upon the zone of delivery and the overall weight of the product.
Using this we are trying to see upon which Order No. we are charged more or lss then the actuall bill price and where is the error.

Atlast we will create a summary of overall overcharged, undercharged and rightfully charged products and their cummulative price and number.

## Dataset Information

We have a total of 5 input files and 1 result file in .xlsx format. 
>Those files and their information are as follows:
* Company X - Order Report : This file conatains ExternOrderNo,SKU and Order Qty details.

* Company X - Pincode Zones : This file conatains Warehouse Pincode, Customer Pincode and Zone details.

* Company X - SKU Master : This file contains SKU and Weight (g) details.

* Courier Company - Invoice : This file contains AWB Code, Order ID, Charged Weight, Warehouse Pincode, Customer Pincode, Zone,Type of Shipment and	Billing Amount (Rs.) details.

* Courier Company - Rates : This file contains fwd_a_fixed, fwd_a_additional, fwd_b_fixed, fwd_b_additional, fwd_c_fixed, fwd_c_additional, fwd_d_fixed, fwd_d_additional, fwd_e_fixed, fwd_e_additional, rto_a_fixed, rto_a_additional, rto_b_fixed, rto_b_additional, rto_c_fixed, rto_c_additional, rto_d_fixed,rto_d_additional,rto_e_fixed and rto_e_additional details.

* Expected_Result : Contains a sample of result generated using python file.

## Guide for running the Code

First install all the python libraries using pip.
```python
pip install -r requirement.txt
```
After importing all the libraries, run the COINTAB_CODE.py code and this will generate a excel file with two sheets.

* Calculation Sheet
* Summary Sheet




