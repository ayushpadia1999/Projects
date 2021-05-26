It is a python desktop app that can calculate the price of Motors based on various
specifications provided by the user and generate a quotation.

The Application uses a python module named Tkinter for the development of the
Graphical User Interface of the application. The Application uses a CSV file to fetch
data regarding the prices of the different standard products sold by the company. The
CSV can be Easily updated by the Non-Technical Employees at the Company who are
the sole users of the app whenever new price sheets are updated. Pandas is used to
fetch the price of the product to be quoted based on the various specifications such as
HP ,KW, IE Rating Provided by the user. Further based on some rules and
specifications the final price of the product is calculated . The user then enters the
discount to be offered and adds it to the list of items to be Quoted which can be seen in
a Text box in the application.

Then when all the products to be quoted have been added to the table. The user
presses the generate Quote button which generates a quotation using the python docx
module to generate the quotation based on the template provided by the company. The
Quotation is saved in a Word file and then later saved with all the other quotations in a
different word file for the record. Logs are maintained in a txt file for easier access to the
data.
