#ASAM: Automated Social Accounting Matrix manipulation
http://www.ecsinsights.org/asam

ASAM is an Excel Application that facilitates the analysis of a sector's contribution to a regional economy by taking IxI data and calculating base economic output, employment, wages and value added. The code (written in Visual Basic) driving the application is Open Source, licensed GNUv3
* the application provides Gross and Base calculations for Output, Employment, Wages and Value Added
* the calculations are automated using Visual Basic
* all calculations are traecible and transparently reflected in the worksheets
* the application makes data available in user-friendly summary tables, pie charts for both Gross and Base values  and bar charts comparing Gross with (directand indirect) Base values.

##Overview
Economists often talk about the "multiplying effect" of a business or sector, e.g: the economic ripple effect of a primary industry by generating incremental demand for supporting businesses as well as induced demand (through its employees) for local retail, hospitality, real estate and other services. It is easy to double count that effect somewhere, making the "exaggeration" a common and fair criticism of economic impact studies. ASAM is an excel application that takes economic sector data (Industry x Industry) and applies a series of "squared" calculations to provide insight in how economic impact cascades through a specific region while keeping the total economic impact of a region whole. This is a well researched methodology (Watson, Philip, Stephen Cooke, David Kay, and Greg Alward. “A Method for Improving Economic Contribution Studies for Regional Analysis.” [Journal of Regional Analysis and Policy, 2015](http://www.researchgate.net/profile/Philip_Watson5/publication/280717696_A_Method_for_Improving_Economic_Contribution_Studies_for_Regional_Analysis/links/55d4b07808ae6788fa352310.pdf)) and the ASAM application merely automates the many steps to achieve the result. By doing so, it significantly reduces the effort and thus barrier in applying the methodology (for example, creating the "square" matrices requires the use of pivoting techniques in MS Access and MS Excel that usually exceeds the know-how of  its users. See also: Doug Olson, "IMPLAN: Excel Pivot to Matrix", IMPLAN Group LLC, 2011. Available at: https://www.youtube.com/watch?v=AVBp9Rbiek8.

Like with any analysis, the quality of results relies on reliable data - garbage in, garbage out. The application will run with any regional IxI sector data. Since IMPLAN is the more widely used data tool for calculating economic contribution in literature it specifically facilitates importing data from IMPLAN's impdb (and earlier revision iap) files.

==============
*An illustrative example, references, strengths and limitation of the method and more information can be found on http://www.ecsinsights.org/asam*


##Installation
You need to have Excel installed on your computer. Note that the calculations in the application are governed by software code which can only execute when “content (macros) are enabled“. If you see a Security Warning you will need to “enable content” or macros.

NOTE: if you do not see this security warning, then your security settings disables macros without notification. You will need to change the setting in your Trust Center to “Disable all macros with notification”. See Microsoft’s instructions for recent Excel versions: https://support.office.com/ or Excel 2003 and older: www.mdmproofing.com/iym/macros.php.

###Basic Use
This 2011 [linked article](http://www.joe.org/joe/2011august/iw3.php) in the *Journal of Extension* provides a compact description of use. The application file itself has a and step-by-step user interface and a FAQ section.

###About the code
It’s not pretty. Most of it is spaghetti code from the initial design (I often just used the VB recorder to automate some calculations) loosely organized in VBA modules. I added an export class to separate the Visual Basic BAS modules from the spreadsheet in order to make version control possible in GIT, but that’s about all the prettifying you’ll find. Any improvements are welcome!

##License
ASAM is released under GNUv3

##Links
The official site for the library is at http://www.ecsinsights.org/asam
