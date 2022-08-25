# Excel-Reverse-Complement
A simple add-in for Excel supplying functions to calculate the reverse, complement, and reverse-complement of a DNA or RNA sequence.

## Installation

Refer to Excel documentation on how to install Excel Add-Ins.  In Excel 2011 for Mac, the procedure is as follows:

 * Select 'Tools->Excel Add-Ins...' and click the 'Select' button
 * Navigate to the .xlam file provided by this package and click 'Open'
 * Ensure that the add-in has a check mark next to it and click 'OK'

## Usage

 ### revcom
 
 This returns the reverse complement of a DNA or RNA sequence. It takes in a string and an optional second parameter specifying whether the string is RNA or DNA:

 =revcom("DNA/RNA SEQUENCE", isRNA)

 replacing "DNA/RNA_SEQUENCE" with the actual sequence or cell reference to be reverse-complemented and isRNA as a 1 if the input sequence is RNA.

 Example:

 =revcom("ATATCGA") will output "TCGATAT"


 ### complement

 This returns the complement of a DNA or RNA sequence. It takes in a string and an optional second parameter specifying whether the string is RNA or DNA:

 =complement("DNA/RNA SEQUENCE", isRNA)

 replacing "DNA/RNA_SEQUENCE" with the actual sequence or cell reference to be complemented and isRNA as a 1 if the input sequence is RNA.

 Example:

 =complement("ATATCGA") will output "TATAGCT"


 ### reverse
 
 This returns the reverse of a string.

 =reverse("String")

 replacing "String" with the actual string to be reversed or a cell reference.

 Example:

 =reverse("ATATCGA") will output "AGCTATA"
 
## License
This code is distributed under the [GNU GPLv3 License](LICENSE).