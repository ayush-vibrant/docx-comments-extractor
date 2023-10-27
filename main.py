import zipfile
from lxml import etree
import argparse
import json

# Namespace for parsing the DOCX XML structure
ooXMLns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def extract_comments_and_associated_text_with_relationships_v2(docxFileName):
    """
    Extracts comments, their associated text, and their relationships (parent-child) from a DOCX file.

    Parameters:
    - docxFileName (str): Path to the DOCX file to be processed.

    Returns:
    - list[dict]: List of dictionaries, where each dictionary contains the associated text and its comments.
    """
    comments_data = {}

    with zipfile.ZipFile(docxFileName) as docx_zip:
        comments_xml = docx_zip.read('word/comments.xml')
        document_xml = docx_zip.read('word/document.xml')

        et_comments = etree.XML(comments_xml)
        et_document = etree.XML(document_xml)

        # Extract all comment elements and their associated ranges in the document
        comments = et_comments.xpath('//w:comment', namespaces=ooXMLns)
        comment_ranges = et_document.xpath('//w:commentRangeStart', namespaces=ooXMLns)

        for comment_range in comment_ranges:
            comment_id = comment_range.xpath('@w:id', namespaces=ooXMLns)[0]

            # Extract the associated text from the main document
            associated_parts = et_document.xpath(
                "//w:r[preceding-sibling::w:commentRangeStart[@w:id=" + comment_id + "] and following-sibling::w:commentRangeEnd[@w:id=" + comment_id + "]]",
                namespaces=ooXMLns
            )
            associated_text = ''.join([part.xpath('string(.)', namespaces=ooXMLns) for part in associated_parts])

            # Extract the comment(s) associated with this text
            comment_texts = [comment.xpath('string(.)', namespaces=ooXMLns) for comment in comments if
                             comment.xpath('@w:id', namespaces=ooXMLns)[0] == comment_id]

            # Group comments by associated text
            if associated_text not in comments_data:
                comments_data[associated_text] = []
            comments_data[associated_text].extend(comment_texts)

    # Convert dictionary to a list for uniformity
    structured_data = [{'associated_text': key, 'comments': value} for key, value in comments_data.items()]

    return structured_data


def parse_command_line_args():
    """
    Parse and return command-line arguments using argparse.

    Returns:
    - argparse.Namespace: Parsed command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description="Extract and display comments and their associated text from a DOCX file.")
    parser.add_argument("filepath", help="Path to the DOCX file to be processed.")
    parser.add_argument("--display-format", choices=["relationship", "original", "enhanced", "json"],
                        default="original",
                        help="Format for displaying the extracted comments. Default is 'original'.")
    parser.add_argument("--include-details", action="store_true",
                        help="If set, includes the author's name and timestamp alongside the comment. Default is to exclude these details.")

    return parser.parse_args()


def display_original_format(data, include_details):
    """Display the extracted data in the original textual format."""
    for i, entry in enumerate(data, 1):
        print(f"Text Segment {i}: {entry['associated_text']}")
        for comment in entry['comments']:
            comment_info = f"  Comment: {comment}"
            print(comment_info)
        print("-" * 50)


def display_enhanced_format(data):
    """Display the extracted data in an enhanced textual format."""
    print("\n" + "=" * 50)
    print("Extracted Comments and Associated Text".center(50))
    print("=" * 50 + "\n")
    for i, entry in enumerate(data, 1):
        print(f"Text Segment {i}:\n{'-' * 15}\n\"{entry['associated_text']}\"\n{'-' * 15}")
        for comment in entry['comments']:
            print(f"Comment: {comment}")
        print("-" * 50 + "\n")


def display_json_format(data):
    """Display the extracted data in JSON format."""
    print(json.dumps(data, indent=4))


def main_cli():
    """
    Main function for command-line execution.
    It processes the provided DOCX file, extracts comments, their associated text, and relationships,
    and then displays the results based on the specified format.
    """
    args = parse_command_line_args()

    if args.display_format == "relationship":
        data = extract_comments_and_associated_text_with_relationships_v2(args.filepath)
        for i, entry in enumerate(data, 1):
            print(f"Text Segment {i}: {entry['associated_text']}")
            if len(entry['comments']) == 1:
                print(f"  Comment: {entry['comments'][0]}")
            else:
                print(f"  Parent Comment: {entry['comments'][0]}")
                for child_comment in entry['comments'][1:]:
                    print(f"    |- Child Comment: {child_comment}")
            print("-" * 50)
    elif args.display_format == "original":
        data = extract_comments_and_associated_text_with_relationships_v2(args.filepath)
        display_original_format(data, args.include_details)
    elif args.display_format == "enhanced":
        data = extract_comments_and_associated_text_with_relationships_v2(args.filepath)
        display_enhanced_format(data)
    elif args.display_format == "json":
        data = extract_comments_and_associated_text_with_relationships_v2(args.filepath)
        display_json_format(data)


if __name__ == "__main__":
    main_cli()
