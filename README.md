# Excel-Reverse-Complement
A simple add-in for Excel supplying one function to calculate the reverse complement of a DNA sequence

## Installation

Refer to Excel documentation on how to install Excel Add-Ins.  In Excel 2011 for Mac, the procedure is as follows:

 * Select 'Tools->Add-Ins...' and click the 'Select' button
 * Navigate to the .xlam file provided by this package and click 'Open'
 * Ensure that the add-in has a check mark next to it and click 'OK'

 ## Usage

 The package provides one function the will calculate the reverse complement of a DNA/RNA sequence and it is called as follows:

 =revcom("DNA_SEQUENCE")

 replacing "DNA_SEQUENCE" with the actual sequence or cell reference to be reverse-complemented.

 Example:

=revcom("ATATCGA") will output "TCGATAT"