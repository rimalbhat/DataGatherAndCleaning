# Data Gathering and Processing

This is a small sample project which gathers Refinery and Blender net input of Crude Oil from [eia.gov](https://www.eia.gov/dnav/pet/pet_pnp_inpt_a_epc0_yir_mbbl_m.htm). The data units of downloaded data is monthly thousand barrels per day. Then it filters that data according to the following conditions.

## How to run the project locally

- **Pre-requisite : python and pip**
- Clone the project and open a termical window in that folder.
- Run the command : 
~~~
pip install -r requirements.txt
~~~
- After that, to run the program, run:
~~~
py app.py
~~~
- Note: if py doesn't work, try python or python3. It depends upon the python installation.
- This will introduce two files into the folder: data.xls, whcich is downloaded from the link and dataProcessed.xlsx, which is the required clean data. 
- It has three sheets, each corresponding to guidelines D, E and F

