# Foreign-Currency-Exchange-UDF-VBA-
Translate Fx rates through a user defined function in excel. The three variable inputs required are the Date of Translation in US Data Format, Base Currency (EX: USD), and Exchanged Currency (EX: GBP).

Steps to implement in excel file:
(1) Create a new module under the personal work book tab, and copy and paste both the function script, and other script contained within. The function script will make the function available, the second script will make the function easily accessible under the "Financial" function category, and will provide descriptions. 
(2) Open the function menu and click under "Financial", find the currencytranslate function and enter in the variable as you would any other excel function.


Please note the function will only work for newer versions of excels that have the webservice built-in function capability. Please feel free to reach out if there are any requests to make this available for older versions of excel. All raw json API data is obtained from http://fixer.io/. Please refer to the fixer.io github here: https://github.com/hakanensari/fixer-io
