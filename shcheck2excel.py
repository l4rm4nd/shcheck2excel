#!/usr/bin/env python
"""
# shcheck2excel
Python script that takes shcheck json and brings the results to Excel

This script will parse json results generated with shcheck and convert them into an Excel file.

The results can be generated with the following command:

    shcheck.py -d -g -j -k --hfile webservices.txt > header_check_new.json

The script can be started with the following command:

    python shcheck2excel -i shcheck_results.json

The output file will be an Excel file showing the results of the audit.
"""

import argparse
import xlsxwriter
import json
from termcolor import colored
import sys
import os

# Write argparser function to parse the following command line arguments
# --inputFile -i (mandatory)
# --outputFile -o (optional)
# --help -h (optional)

def parseArgs():
    banner = """
    

       .__           .__                   __   ________                           .__   
  _____|  |__   ____ |  |__   ____   ____ |  | _\_____  \ ____ ___  ___ ____  ____ |  |  
 /  ___/  |  \_/ ___\|  |  \_/ __ \_/ ___\|  |/ //  ____// __ \\  \/  // ___\/ __ \|  |  
 \___ \|   Y  \  \___|   Y  \  ___/\  \___|    </       \  ___/ >    <\  \__\  ___/|  |__
/____  >___|  /\___  >___|  /\___  >\___  >__|_ \_______ \___  >__/\_ \\___  >___  >____/
     \/     \/     \/     \/     \/     \/     \/       \/   \/      \/    \/    \/      


    """

    parser = argparse.ArgumentParser(description='shcheck2excel', epilog=banner, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-i', '--inputFile', help='shcheck results .json format', required=True)
    parser.add_argument('-o', '--outputFile', help='Excel file to write the results to', required=False)

    args = parser.parse_args()
    return args


############################################
#### Classes
############################################

class ResponseHeaderAnalysisResult:
    def __init__(self, url, x_xss_protection, x_frame_options, x_content_type_options, content_security_policy, x_permitted_cross_domain_policies, referrer_policy, expect_ct, permissions_policy, cross_origin_embedder_policy, cross_origin_resource_policy, cross_origin_opener_policy, feature_policy, strict_transport_security):
        self.url = url
        self.x_xss_protection = x_xss_protection
        self.x_frame_options = x_frame_options
        self.x_content_type_options = x_content_type_options
        self.content_security_policy = content_security_policy
        self.x_permitted_cross_domain_policies = x_permitted_cross_domain_policies
        self.referrer_policy = referrer_policy
        self.expect_ct = expect_ct
        self.permissions_policy = permissions_policy
        self.cross_origin_embedder_policy = cross_origin_embedder_policy
        self.cross_origin_resource_policy = cross_origin_resource_policy
        self.cross_origin_opener_policy = cross_origin_opener_policy
        self.feature_policy = feature_policy
        self.strict_transport_security = strict_transport_security


############################################
#### Functions
############################################
# Generate an function that takes the ResponseHeaderAnalysisResultsArr and renders the results in Excel
def generateReport(ResponseHeaderAnalysisResultsArr, outputFile):
    
    # get current path
    currentPath = os.path.dirname(os.path.realpath(__file__))
    # echo Writing results to Excel file path of the outputFile
    print(colored("Writing results to Excel file: " + currentPath + "/" + outputFile, 'green'))

    # Create a workbook
    workbook = xlsxwriter.Workbook(outputFile)

    # Add a worksheet
    worksheet = workbook.add_worksheet("Overview")

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some data headers.
    worksheet.write('B1', 'X-XSS-Protection header', bold)
    worksheet.write('C1', 'X-Frame-Options header', bold)
    worksheet.write('D1', 'X-Content-Type-Options header', bold)
    worksheet.write('E1', 'Content-Security-Policy header', bold)
    worksheet.write('F1', 'X-Permitted-Cross-Domain-Policies header', bold)
    worksheet.write('G1', 'Referrer-Policy header', bold)
    worksheet.write('H1', 'Strict-Transport-Security', bold)
    worksheet.write('I1', 'Permissions-Policy header', bold)
    worksheet.write('J1', 'Cross-Origin-Embedder-Policy header', bold)
    worksheet.write('K1', 'Cross-Origin-Resource-Policy header', bold)
    worksheet.write('L1', 'Cross-Origin-Opener-Policy header', bold)
    worksheet.write('M1', 'Feature-Policy header', bold)
    worksheet.write('N1', 'Expect-CT header', bold)

    # Set a filter on the first column to the last column
    worksheet.autofilter('A1:N1')

    # Start from the first cell below the headers.
    row = 1
    col = 0

    # Iterate over the data and write it out row by row.
    for result in ResponseHeaderAnalysisResultsArr:
        worksheet.write(row, col, result.url)
        # Write the present headers of the x_xss_protection object
        # if the header is present format the cell green
        # if the header is not present format the cell red
        if result.x_xss_protection["present"] == True:
            worksheet.write(row, col + 1, result.x_xss_protection["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 1, result.x_xss_protection["present"], workbook.add_format({'bg_color': 'red'}))
        
        if result.x_frame_options["present"] == True:
            worksheet.write(row, col + 2, result.x_frame_options["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 2, result.x_frame_options["present"], workbook.add_format({'bg_color': 'red'}))

        if result.x_content_type_options["present"] == True:
            worksheet.write(row, col + 3, result.x_content_type_options["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 3, result.x_content_type_options["present"], workbook.add_format({'bg_color': 'red'}))
        
        if result.content_security_policy["present"] == True:
            worksheet.write(row, col + 4, result.content_security_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 4, result.content_security_policy["present"], workbook.add_format({'bg_color': 'red'}))

        if result.x_permitted_cross_domain_policies["present"] == True:
            worksheet.write(row, col + 5, result.x_permitted_cross_domain_policies["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 5, result.x_permitted_cross_domain_policies["present"], workbook.add_format({'bg_color': 'red'}))

        if result.referrer_policy["present"] == True:
            worksheet.write(row, col + 6, result.referrer_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 6, result.referrer_policy["present"], workbook.add_format({'bg_color': 'red'}))

        if result.strict_transport_security["present"] == True:
            worksheet.write(row, col + 7, result.strict_transport_security["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 7, result.strict_transport_security["present"], workbook.add_format({'bg_color': 'red'}))

        if result.permissions_policy["present"] == True:
            worksheet.write(row, col + 8, result.permissions_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 8, result.permissions_policy["present"], workbook.add_format({'bg_color': 'red'}))

        if result.cross_origin_embedder_policy["present"] == True:
            worksheet.write(row, col + 9, result.cross_origin_embedder_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 9, result.cross_origin_embedder_policy["present"], workbook.add_format({'bg_color': 'red'}))

        if result.cross_origin_resource_policy["present"] == True:
            worksheet.write(row, col + 10, result.cross_origin_resource_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 10, result.cross_origin_resource_policy["present"], workbook.add_format({'bg_color': 'red'}))

        if result.cross_origin_opener_policy["present"] == True:
            worksheet.write(row, col + 11, result.cross_origin_opener_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 11, result.cross_origin_opener_policy["present"], workbook.add_format({'bg_color': 'red'}))

        if result.feature_policy["present"] == True:
            worksheet.write(row, col + 12, result.feature_policy["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 12, result.feature_policy["present"], workbook.add_format({'bg_color': 'red'}))
        
        if result.expect_ct["present"] == True:
            worksheet.write(row, col + 13, result.expect_ct["present"], workbook.add_format({'bg_color': 'green'}))
        else:
            worksheet.write(row, col + 13, result.expect_ct["present"], workbook.add_format({'bg_color': 'red'}))

        row += 1
    

    ############################
    ## Present Headers Sheet ##
    ############################

    # Add a new worksheet to the workbook
    worksheet = workbook.add_worksheet()
    worksheet.name = "Present Headers"

    # Set a filter on the first column
    worksheet.autofilter('A1:C1')

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # First culumn is the URL
    # Second column is the Header
    # Third column is the Value
    worksheet.write('A1', 'URL', bold)
    worksheet.write('B1', 'Header', bold)
    worksheet.write('C1', 'Value', bold)

    row = 1
    col = 0



    for result in ResponseHeaderAnalysisResultsArr:
        
        if result.x_xss_protection["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "X-XSS-Protection")
            worksheet.write(row, col + 2, result.x_xss_protection["value"])
            row += 1

        if result.x_frame_options["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "X-Frame-Options")
            worksheet.write(row, col + 2, result.x_frame_options["value"])
            row += 1

        if result.x_content_type_options["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "X-Content-Type-Options")
            worksheet.write(row, col + 2, result.x_content_type_options["value"])
            row += 1

        if result.content_security_policy["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Content-Security-Policy")
            worksheet.write(row, col + 2, result.content_security_policy["value"])
            row += 1

        if result.x_permitted_cross_domain_policies["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "X-Permitted-Cross-Domain-Policies")
            worksheet.write(row, col + 2, result.x_permitted_cross_domain_policies["value"])
            row += 1

        if result.referrer_policy["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Referrer-Policy")
            worksheet.write(row, col + 2, result.referrer_policy["value"])
            row += 1
        
        if result.strict_transport_security["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Strict-Transport-Security")
            worksheet.write(row, col + 2, result.strict_transport_security["value"])
            row += 1
        
        if result.cross_origin_embedder_policy["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Cross-Origin-Embedder-Policy")
            worksheet.write(row, col + 2, result.cross_origin_embedder_policy["value"])
            row += 1
        
        if result.cross_origin_resource_policy["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Cross-Origin-Resource-Policy")
            worksheet.write(row, col + 2, result.cross_origin_resource_policy["value"])
            row += 1
        
        if result.cross_origin_opener_policy["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Cross-Origin-Opener-Policy")
            worksheet.write(row, col + 2, result.cross_origin_opener_policy["value"])
            row += 1

        if result.feature_policy["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Feature-Policy")
            worksheet.write(row, col + 2, result.feature_policy["value"])
            row += 1

        if result.expect_ct["present"] == True:
            worksheet.write(row, col, result.url)
            worksheet.write(row, col + 1, "Expect-CT")
            worksheet.write(row, col + 2, result.expect_ct["value"])
            row += 1

    workbook.close()

def main():
    # get arguments
    args = parseArgs()

    # get the JSON file
    jsonFile = args.inputFile

    # check if the JSON file exists
    if not os.path.exists(jsonFile):
        print(f'JSON file {jsonFile} does not exist')
        sys.exit(1)
    
    # check if output file is set
    if args.outputFile is None:
        outputFile = 'shcheck2excel-results.xlsx'
    else:
        outputFile = args.outputFile
    
    # Open JSON file
    results = open(jsonFile, 'r')

    # Load JSON file
    headerAnalysisResults = json.load(results)

    # Close JSON file
    results.close()

    # Create an array to hold the ResponseHeaderAnalysisResult objects
    ResponseHeaderAnalysisResultsArr = []

    # iterate over URLs
    for url, headers in headerAnalysisResults.items():
        # print the URL
        print(url)
        print (colored('   Present Headers:', 'green'))
        for key, value in headers['present'].items():
            # print the header and its value
            print(f'       {key}: {value}')

        # iterate over missing headers
        print (colored('   Missing Headers:', 'red'))
        for header in headers['missing']:
            # print the header
            print(f'       {header}')
        print('\n') 

        # Create a ResponseHeaderAnalysisResult object
        x_xss_protection = {"present": False, "value": ""}
        x_frame_options = {"present": False, "value": ""}
        x_content_type_options = {"present": False, "value": ""}
        content_security_policy = {"present": False, "value": ""}
        x_permitted_cross_domain_policies = {"present": False, "value": ""}
        referrer_policy = {"present": False, "value": ""}
        expect_ct = {"present": False, "value": ""}
        permissions_policy = {"present": False, "value": ""}
        cross_origin_embedder_policy = {"present": False, "value": ""}
        cross_origin_resource_policy = {"present": False, "value": ""}
        cross_origin_opener_policy = {"present": False, "value": ""}
        feature_policy = {"present": False, "value": ""}
        strict_transport_security = {"present": False, "value": ""}

        # iterate over present headers
        for key, value in headers['present'].items():
            # Check if key is "X-XSS-Protection"
            if key == "X-XSS-Protection":
                x_xss_protection = {"present": True, "value": value}
            
            # Check if key is "X-Frame-Options"
            if key == "X-Frame-Options":
                x_frame_options = {"present": True, "value": value}

            # Check if key is "X-Content-Type-Options"
            if key == "X-Content-Type-Options":
                x_content_type_options = {"present": True, "value": value}

            # Check if key is "Content-Security-Policy"
            if key == "Content-Security-Policy":
                content_security_policy = {"present": True, "value": value}

            # Check if key is "X-Permitted-Cross-Domain-Policies"
            if key == "X-Permitted-Cross-Domain-Policies":
                x_permitted_cross_domain_policies = {"present": True, "value": value}

            # Check if key is "Referrer-Policy"
            if key == "Referrer-Policy":
                referrer_policy = {"present": True, "value": value}

            # Check if key is "Expect-CT"
            if key == "Expect-CT":
                expect_ct = {"present": True, "value": value}

            # Check if key is "Permissions-Policy"
            if key == "Permissions-Policy":
                permissions_policy = {"present": True, "value": value}

            # Check if key is "Cross-Origin-Embedder-Policy"
            if key == "Cross-Origin-Embedder-Policy":
                cross_origin_embedder_policy = {"present": True, "value": value}

            # Check if key is "Cross-Origin-Resource-Policy"
            if key == "Cross-Origin-Resource-Policy":
                cross_origin_resource_policy = {"present": True, "value": value}

            # Check if key is "Cross-Origin-Opener-Policy"
            if key == "Cross-Origin-Opener-Policy":
                cross_origin_opener_policy = {"present": True, "value": value}

            # Check if key is "Feature-Policy"
            if key == "Feature-Policy":
                feature_policy = {"present": True, "value": value}

            # Check if key is "Strict-Transport-Security"
            if key == "Strict-Transport-Security":
                strict_transport_security = {"present": True, "value": value}

        # Create a ResponseHeaderAnalysisResult object
        analysisResult=ResponseHeaderAnalysisResult(url, x_xss_protection, x_frame_options, x_content_type_options, content_security_policy, x_permitted_cross_domain_policies, referrer_policy, expect_ct, permissions_policy, cross_origin_embedder_policy, cross_origin_resource_policy, cross_origin_opener_policy, feature_policy, strict_transport_security)

        # Add the ResponseHeaderAnalysisResult object to the array
        ResponseHeaderAnalysisResultsArr.append(analysisResult)    

    # Generate report
    generateReport(ResponseHeaderAnalysisResultsArr, outputFile)

if __name__ == "__main__":
    main()
