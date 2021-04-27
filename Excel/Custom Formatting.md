# Custom Formatting
There are many useful formats in Excel, but you can set it in a custom way yourself using a small code. This guide gives a few that can be useful.

## Structure and Reference:
Excel custom number formats have a specific structure. Each number format can have up to four sections, separated with semi-colons as follows:
1;2;3;4@
This structure can make custom number formats look overwhelmingly complex. To read a custom number format, learn to spot the semi-colons and mentally parse the code into these sections:
1. Positive values
2. Negative values
3. Zero values
4. Text values

## Not all sections required:
Although a number format can include up to four sections, only one section is required. By default, the first section applies to positive numbers, the second section applies to negative numbers, the third section applies to zero values, and the fourth section applies to text.
When only one format is provided, Excel will use that format for all values.
If you provide a number format with just two sections, the first section is used for positive numbers and zeros, and the second section is used for negative numbers.
To skip a section, include a semi-colon in the proper location, but don't specify a format code.

## Characters that display natively:
Some characters appear normally in a number format, while others require special handling. The following characters can be used without any special handling:
| Character | Comment |
| :---: | :---: |
| $ | Dollar |
| +- | Plus, Minus |
| () | Parentheses |
| {} | Curly braces |
| <> | Less than, greater than |
| = | Equal |
| : | Colon |
| ^ | Caret |
| ' | Apostrophe |
| / | Forward Slash |
| ! | Exclamation point |
| & | Ampersand |
| ~ | Tilde |
|   | Space Character |

## Escaping characters:
Some characters won't work correctly in a custom number format without being escaped. For example, the asterisk (*), hash (#), and percent (%) characters can't be used directly in a custom number format – they won't appear in the result. The escape character in custom number formats is the backslash (\ ). By placing the backslash before the character, you can use them in custom number formats:
| Value | Code | Result |
| :---: | :---: | :---: |
| 100 | \#0 | #100 |
| 100 | \*0 | *100 |
| 100 | \%0 | %100 |

## Placeholders
Certain characters have special meaning in custom number format codes. The following characters are key building blocks:
| Character | Purpose |
| :---: | :---: |
| 0 | Display insignificant zeros |
| # |	Display significant digits |
| ? |	Display aligned decimals |
| . |	Decimal point |
| , |	Thousands separator |
| * |	Repeat digit |
| _ |	Add space |
| @ |	Placeholder for text |
- Zero (0) is used to force the display of insignificant zeros when a number has fewer digits than than zeros in the format. For example, the custom - format 0.00 will display zero as 0.00, 1.1 as 1.10 and .5 as 0.50.
- Pound sign (#) is a placeholder for optional digits. When a number has fewer digits than # symbols in the format, nothing will be displayed. For example, the custom format #.## will display 1.15 as 1.15 and 1.1 as 1.1.
- Question mark (?) is used to align digits. When a question mark occupies a place not needed in a number, a space will be added to maintain visual alignment.
- Period (.) is a placeholder for the decimal point in a number. When a period is used in a custom number format, it will always be displayed, regardless of whether the number contains decimal values.
- Comma (,) is a placeholder for the thousands separators in the number being displayed.  It can be used to define the behavior of digits in relation to the thousands or millions digits.
- Asterisk (*) is used to repeat characters. The character immediately following an asterisk will be repeated to fill remaining space in a cell.
- Underscore (_ ) is used to add space in a number format. The character immediately following an underscore character controls how much space to add. A common use of the underscore character is to add space to align positive and negative values when a number format is adding parentheses to negative numbers only. For example, the number format "0_);(0)" is adding a bit of space to the right of positive numbers so that they stay aligned with negative numbers, which are enclosed in parentheses.
- At (@) - placeholder for text. For example, the following number format will display text values in blue:

# Template
Code: 000
## Explanation:
This is a template for how each item should be laid out.
## Example (Input = Output):
Input Value = Output Value Examples (Positive, Negative, 0s)
## Formula:
 FORMULA

# Commas and Blank 0s
Code: 001
## Explanation:
This format will add thousand separators (ie. 1,000), and blank out any values of 0
## Example (Input = Output):
1001 = 1,001, -1001 = -1,001, 0 = [BLANK]
## Formula:
 #,##0;-#,##0;;@

# Blank 0s Accounting
Code: 002
## Explanation:
This format will add thousand separators (ie. 1,000), put negative values into brackets, and blank out any values of 0
## Example (Input = Output):
1000 = £1,000, -1000 = £(1000), 0 = [BLANK]
## Formula:
 £#,##0;£(#,##0);;@
