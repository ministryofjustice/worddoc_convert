import os
import urllib2
import csv
import shutil
from time import sleep
from optparse import OptionParser

def file_exists(file_path):
    try:
        with open(file_path) as f: pass
    except IOError as e:
        return False
    return True

def remove_if_exists(file_path):
    try:
        os.remove(file_path)
    except OSError:
        pass

def download_files(csv_file_path, docs_output_dir):

    #create docs directory
    os.makedirs(docs_output_dir)

    with open(csv_file_path, 'rb') as csv_file:
        dialect = csv.Sniffer().sniff(csv_file.read(1024))
        csv_file.seek(0)
        reader = csv.reader(csv_file)
        for row in reader:
            try:
                url = row[0].strip()
                content = urllib2.urlopen(url).read()
                file_name = url.split('/')[len(url.split('/')) - 1]
                file_path = '%s/%s' % (docs_output_dir, file_name)
                remove_if_exists(file_path)
                doc = open(file_path, "wb")
                doc.write(content)
                doc.close()
                print "Downloaded %s" % file_name
                sleep(0.5)
            except:
                print "FAIL: Could not download file %s " % row[0]

def convert_files(input_dir, output_dir, openoffice_cli):

    #create directories (todo: rm -r if exists)
    for format in  ['html', 'pdf',]:
        if not os.path.exists("%s/%s" % (output_dir, format)):
            os.makedirs("%s/%s" % (output_dir, format))

    # convert file
    for doc in os.listdir(input_dir):
        if doc.lower().endswith(".doc") or doc.lower().endswith(".rtf"):
            try:
                file_path ='%s/%s' % (input_dir, doc)
                for format in  ['html', 'pdf',]:
                    command =  '%s --headless --convert-to %s %s --outdir %s' % (openoffice_cli, format, file_path, output_dir + '/' + format)
                    os.system(command)
            except:
                print "FAIL: Could not convert %s" % doc

parser = OptionParser()
parser.add_option("-i", "--input", dest="input",
                  help="CSV file containing list of urls of word documents")
parser.add_option("-o", "--output", dest="output",
                  help="Directory to output things to")
parser.add_option("-c", "--openofficecli", dest="openofficecli", default = '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                  help="Location of open office / libra office cli")

(options, args) = parser.parse_args()

# check input file
if not options.input:
    parser.error('You must supply an input csv file')
else:
    if not file_exists(options.input):
        parser.error('%s does not exist' % options.input)

# check output dir and create if needed
if not options.output:
    parser.error('You must supply an output directory')
else:
    output_dir = options.output + '/worddoc_convert_output'
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

#check for open office
if not file_exists(options.openofficecli):
    parser.error('Could not find open office. Make sure you have it installed')

download_files(options.input, output_dir + '/docs')
convert_files(output_dir + '/docs', output_dir, options.openofficecli)
