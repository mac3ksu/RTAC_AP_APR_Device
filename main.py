import xml.etree.ElementTree as ET
import xlrd
import os


def build_apr(port, name):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_xml = os.path.join(dir_path, 'Templates', 'APR_Template.xml')
    apr_name = 'APR_' + name
    device_name = name
    port = port
    outfile_name = 'APR_' + name + '.xml'
    outfile_xml = os.path.join(dir_path, 'Output Files', outfile_name)

    with open(template_xml, 'r') as infile:
        with open(outfile_xml, 'w') as outfile:
            for line in infile:
                line = line.replace('__APR_Name__', apr_name)
                line = line.replace('__Port__', port)
                line = line.replace('__Device_Name__', device_name)
                outfile.write(line)
        outfile.close()
    infile.close()


def build_ap_ssh(port, name):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_xml = os.path.join(dir_path, 'Templates', 'AP_Eth_SSH_Template.xml')
    device_name = name
    port = port
    outfile_name = name + '_AP.xml'
    outfile_xml = os.path.join(dir_path, 'Output Files', outfile_name)

    with open(template_xml, 'r') as infile:
        with open(outfile_xml, 'w') as outfile:
            for line in infile:
                line = line.replace('__Port__', port)
                line = line.replace('__Device_Name__', device_name)
                outfile.write(line)
        outfile.close()
    infile.close()


def build_501(port, name):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_xml = os.path.join(dir_path, 'Templates', '501_Template.xml')
    device_name = name
    port = port
    outfile_name = name + '_SEL.xml'
    outfile_xml = os.path.join(dir_path, 'Output Files', outfile_name)

    with open(template_xml, 'r') as infile:
        with open(outfile_xml, 'w') as outfile:
            for line in infile:
                line = line.replace('__Port__', port)
                line = line.replace('__Device_Name__', device_name)
                outfile.write(line)
        outfile.close()
    infile.close()


def build_ap_satec(port, name):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_xml = os.path.join(dir_path, 'Templates', 'AP_Serial_Satec_Template.xml')
    device_name = name
    port = port
    outfile_name = name + '_AP.xml'
    outfile_xml = os.path.join(dir_path, 'Output Files', outfile_name)

    with open(template_xml, 'r') as infile:
        with open(outfile_xml, 'w') as outfile:
            for line in infile:
                line = line.replace('__Port__', port)
                line = line.replace('__Device_Name__', device_name)
                outfile.write(line)
        outfile.close()
    infile.close()


def build_421(port, name):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    template_xml = os.path.join(dir_path, 'Templates', '421_Template.xml')
    device_name = name
    port = port
    outfile_name = name + '_SEL.xml'
    outfile_xml = os.path.join(dir_path, 'Output Files', outfile_name)

    with open(template_xml, 'r') as infile:
        with open(outfile_xml, 'w') as outfile:
            for line in infile:
                line = line.replace('__Port__', port)
                line = line.replace('__Device_Name__', device_name)
                outfile.write(line)
        outfile.close()
    infile.close()


if __name__ == '__main__':
    wbook = xlrd.open_workbook('Input.xlsx')
    wsheet = wbook.sheet_by_index(0)
    row = 1

    while row <= wsheet.nrows:
        try:
            if wsheet.cell_value(row, 0) != '':
                print(wsheet.cell_value(row,0))
                print(wsheet.cell_value(row,1))
                print(wsheet.cell_value(row,2))
                if wsheet.cell_value(row, 1) == 'APR':
                    build_apr(str(int(wsheet.cell_value(row, 0))), wsheet.cell_value(row, 2))
                if wsheet.cell_value(row,1) == 'AP - Eth (SSH)':
                    build_ap_ssh(str(int(wsheet.cell_value(row, 0))), wsheet.cell_value(row, 2))
                if wsheet.cell_value(row,1) == 'AP - Serial Satec':
                    build_ap_satec(str(int(wsheet.cell_value(row, 0))), wsheet.cell_value(row, 2))
                if str(int(wsheet.cell_value(row,1))) == '501':
                    build_501(str(int(wsheet.cell_value(row, 0))), wsheet.cell_value(row, 2))
                if str(int(wsheet.cell_value(row,1))) == '421':
                    build_421(str(int(wsheet.cell_value(row, 0))), wsheet.cell_value(row, 2))
        except:
            pass
        row = row + 1
