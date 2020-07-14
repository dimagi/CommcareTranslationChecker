import re

INLINE_FORMATTING_TAGS = [
    r'(\*\*[\S]+)',  # opening format tag for bold
    r'([\S]+\*\*)',  # closing format tag for bold
    r'(\*[\S]+)',    # opening format tag for italic
    r'([\S]+\*)',    # closing format tag for italic
    r'(\*\*\*[\S]+)',  # opening format tag for bold italic
    r'([\S]+\*\*\*)',  # closing format tag for bold italic
    r'(~~[\S]+)', # opening format tag for strikethrough
    r'([\S]+~~)', # closing format tag for strikethrough
]
BLOCK_FORMATTING_TAGS = [
    r'(^# [\S]+)', # Format tag for heading 1
    r'(^## [\S]+)', # Format tag for heading 2
    r'(^### [\S]+)', # Format tag for heading 3
    r'(^#### [\S]+)', # Format tag for heading 4
    r'(^##### [\S]+)', # Format tag for heading 5
    r'(^###### [\S]+)', # Format tag for heading 1
    r'(^\* [\S]+)', # Format tag for unordered lists
    r'(^[0-9]+. [\S]+)' # Format tag for ordered lists
]


def regex_match_count(expr, text):
    """
    Check for regular expression match in every line of the text
    and return the count of matches
    """
    count = 0
    text_lines = text.splitlines()
    for line in text_lines:
        if len(re.findall(expr, line)) > 0:
            count += 1
    return count


def normalizeQuotes(text):
    """
    Unicode contains multiple forms of quotes. This method replaces every single quote
    and double quote to their ascii equivalent.
    Converts ‘, ’, ‛, ❛, ❜ to '
    Converts “, ”, ‟, ❝, ❞, ＂ to "
    """
    single_quotes_ord_values = [700, 1370, 8216, 8217, 8219, 10075, 10076, 65287]
    double_quotes_ord_values = [750, 8220, 8221, 8223, 10077, 10078, 65282]

    normalized_str = []
    if not text:
        return ""
    for char in text:
        if ord(char) in single_quotes_ord_values:
            normalized_str.append(chr(39))
        elif ord(char) in double_quotes_ord_values:
            normalized_str.append(chr(34))
        else:
            normalized_str.append(char)
    return ''.join(normalized_str)


def fix_block_tags_mismatch(baseText, outputText):
    """
    Block tags such as headings, lists occur at the start of the line.
    We compare each line in base text with every line in output text.
    If base text has a block tag in a line, then output text must also
    have a block tag in the same line. If block tag not found or some characters
    of block tag is missing, we fix them

    Input:
    baseText(str): base column text
    outputText(str): Comparing column text

    :return: None or Fixed output text
    """
    baseTextLines = baseText.splitlines()
    outputTextLines = outputText.splitlines()
    fixed_output_text = []

    # If line count in base text and output text doesn't match,
    # then we cannot compare base text and output text line by line
    # so we return None which implies we are unable to fix
    # A Warning will still be displayed to the user for this cell
    if len(baseTextLines) != len(outputTextLines):
        return None

    try:
        # Compare each line of baseText with outputText,
        # if mismatch occurs, fix it
        for index in range(len(baseTextLines)):
            baseTextLine = baseTextLines[index]
            outputTextLine = outputTextLines[index]
            regex_match_found = False

            for tag in BLOCK_FORMATTING_TAGS:
                base_tag_char_list = []
                # Find a match for a block tag, if match found in baseText line,
                # for each character in the block tag in baseTextLine, compare with characters in
                # outputTextLine, if all characters in block tag are matched, then we add the line to output
                # otherwise remove the text till mismatch position and prepend with block tag
                if len(re.findall(tag, baseTextLine)) > 0:
                    regex_match_found = True

                    for char in baseTextLine:
                        base_tag_char_list.append(char)
                        if char == ' ':
                            break

                    for position in range(len(base_tag_char_list)):
                        if outputTextLine[position] != base_tag_char_list[position]:
                            break

                    if position == len(base_tag_char_list):
                        fixed_output_text.append(outputTextLine)
                        break
                    else:
                        fixed_output_text.append(re.sub(' +', ' ', ''.join(base_tag_char_list) + outputTextLine[position:]))
                        break

            if regex_match_found is False:
                fixed_output_text.append(outputTextLine)
        return '\n'.join(fixed_output_text)
    except Exception as e:
        # If any exception occurs while trying to fix block tag mismatch, we return None
        # as we are unable to fix mismatches
        print("Exception while fixing block tag mismatch %s" %(e))
        return None


def removeExtraOutputValues(extraValueList, text):
    """
    Removes Output tags with values in extraValueList
    Input:
    extraValueList(list): list of output tag values to be removed
    text(str): The text in which extra output values to be removed

    :return: outputText(str): The text with extra output values removed
    """
    outputText = text
    for extraValue in extraValueList:
        extraTagToRemove = f'<output value="{extraValue}"/>'
        outputText = outputText.replace(extraTagToRemove, '')
    return outputText


def swapOutputValues(outputValueList, text):
    """
    Swap output tags in the text if they are out of order
    Input:
    outputValueList(list): list of output values
    text(str): The text in which output values need to be swapped

    :return: str:The text with output values swapped if len(outpputValueList) is 2
                otherwise return text which is passed as parameter
    """
    if len(outputValueList) == 2:
        tag1 = f'<output value="{outputValueList[0]}"/>'
        tag2 = f'<output value="{outputValueList[1]}"/>'
        tag1Start = text.find(tag1)
        tag1End = tag1Start + len(tag1)
        tag2Start = text.find(tag2)
        tag2End = tag2Start + len(tag2)
        return text[:tag1Start] + tag2 + text[tag1End:tag2Start] + tag1 + text[tag2End:]
    else:
        return text
