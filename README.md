After my first attempt to clean the disorganized customer details dataset, I've been searching for a more straightforward yet effective method to tidy up the messy data using Excel formulas. Finally, my search led me to the solution I was seeking. This article provides a detailed approach to cleaning the dataset using the Excel formula I discovered.

The dataset cleaning process involves the utilization of several Excel functions, including:

1. LET()
2. TEXTAFTER()
3. TEXTBEFORE()
4. LEFT()
5. TRIM()

These functions play a crucial role in achieving an organized and refined dataset.

The disorganized dataset was downloaded from forsightBi using the provided link. It contains customer details that are jumbled up together in a single cell within the spreadsheet.

The information includes customer names, addresses, ages, and genders.


![THE dataset](https://github.com/dannieRope/Cleaning-Jumbled-Up-Customers-Details-in-Excel--APPROACH-1/assets/132214828/4cdf300d-88e8-4d78-a033-a1bda5caf34d)


The primary goal of this data cleaning exercise is to extract and organize the customer details into separate columns, ensuring that each piece of information such as customer names, addresses, ages, and genders occupies its designated column for a well-structured dataset.

DATA CLEANING PROCESS

To extract the customer names, the following formula was used:

=LET(a, TEXTAFTER(A13, "Name"), b, TEXTBEFORE(a, "Address"), b)

This formula uses the LET() function to define variables a and b. The TEXTAFTER() function extracts the text after the word "Name" in cell A13 and assigns it to variable a. Then, the TEXTBEFORE() function extracts the text before the word "Address" from variable a and assigns it to variable b. Finally, the formula returns the value of variable b, which represents the customer name.

![name](https://github.com/dannieRope/Cleaning-Jumbled-Up-Customers-Details-in-Excel--APPROACH-1/assets/132214828/2c705d98-0bae-4c77-8a27-6c5128a49879)



To extract the addresses, the following formula was used:

=LET(a, TEXTBEFORE(A13, "Age"), b, TEXTAFTER(a, "Address"), b)

Similarly, this formula also uses the LET() function to define variables a and b. The TEXTBEFORE() function extracts the text before the word "Age" in cell A13 and assigns it to variable a. Then, the TEXTAFTER() function extracts the text after the word "Address," and assigns it to variable b. Finally, the formula returns the value of variable b, which represents the customer address.

![address](https://github.com/dannieRope/Cleaning-Jumbled-Up-Customers-Details-in-Excel--APPROACH-1/assets/132214828/9f81cc7a-e93e-4dba-a530-3b865a72946c)

To extract the customer age details the following formula was used:

=LET(a, TEXTAFTER(A13, "Age"), b, TRIM(LEFT(a, 3)), b)

As before, this formula uses the LET() function to define variables a and b. The TEXTAFTER() function extracts the text after the word "Age" in cell A13 and assigns it to variable a. Then, the LEFT() function takes the leftmost three characters from variable a, representing the age, and the TRIM() function removes any leading or trailing spaces. The result is then assigned to variable b, and the formula returns the value of variable b, which represents the customer age details.

![age](https://github.com/dannieRope/Cleaning-Jumbled-Up-Customers-Details-in-Excel--APPROACH-1/assets/132214828/90416ef7-d6e0-4eba-bc9c-96e3fd7d0b53)

To extract the gender details, the following formula was used:

=TRIM(TEXTAFTER(A13, "Gender"))

In this case, the TEXTAFTER() function is used to extract the text after the word "Gender" in cell A13, which represents the gender information. The TRIM() function is then used to remove any leading or trailing spaces, ensuring a clean and tidy result. The formula returns the gender details extracted from the cell A13.

![gender](https://github.com/dannieRope/Cleaning-Jumbled-Up-Customers-Details-in-Excel--APPROACH-1/assets/132214828/53870e38-bde5-4235-815b-fd28543bdccf)


After applying these formulas, you will have all the customer information, including names, addresses, ages, and genders, extracted and displayed accurately in their respective columns. The dataset is now cleaned and properly organized, making it ready for further analysis or use. Remember to save your cleaned dataset for future reference, as it now contains relevant information and is easier to work with.


![final](https://github.com/dannieRope/Cleaning-Jumbled-Up-Customers-Details-in-Excel--APPROACH-1/assets/132214828/9e743c53-4b34-4d11-aab2-37471bc5c41a)


Check out the first approach here.
