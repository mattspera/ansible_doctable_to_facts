#!/usr/bin/python

# Copyright: (c) 2018, Matthew Spera <speramatthew@gmail.com>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

ANSIBLE_METADATA = {
    'metadata_version': '1.1',
    'status': ['preview'],
    'supported_by': 'community'
}

DOCUMENTATION = '''
---
module: doctable_to_facts

short_description: Read and parse specific tables within a .docx file and output Ansible facts.

description:
    - Read and parse specific tables within a .docx file and output Ansible facts.
    - Ansible facts are created in the form of a list with each element in the list
      a dictionary. The table column header is used as the key and the contents of
      the cell used as the value. A dictionary is created for each table parsed.
    
requirements:
    - python-docx can be obtained from PyPi (https://pypi.org/project/python-docx)

options:
    src:
        description:
            - The name of the word document.
        required: true
    name:
        description:
            - Name to give as reference to table data in Ansible facts.
        required: true
    headers:
        description:
            - List of table headers to search for to identify desired tables to parse.
        type: list
        required: true

author:
    - Matthew Spera (@smatt241)
'''

EXAMPLES = '''
# Parse specific table within document
- name: Read static route table within document
  doctable_to_facts:
    src: test_document.docx
    name: 'route_table'
    headers: ['Destination', 'NextHop']
	
# Parse multiple tables within document
- name: Read static route and zone tables within document
  doctable_to_facts:
    src: test_document.docx
    name: '{{ item.name }}'
    headers: '{{ item.headers }}'
  with_items:
    - name: routes
      headers: ['Destination', 'NextHop']
    - name: zones
      headers: ['ZoneProtectionProfile']
'''

RETURN = '''
facts_json:
    description: Table data output as Ansible facts.
message:
    description: The output message generated.
'''

import json
from ansible.module_utils.basic import AnsibleModule

try:
    import docx
    
    HAS_LIB = True
except ImportError:
    HAS_LIB = False

def parse_tables_dict(input_file, table_name, headers):
    result = {'ansible_facts':{}}
    tables = {}
    
    ansible_table_name = 'table_' + table_name
    tables[ansible_table_name] = []
    
    try:
        doc = docx.Document(input_file)
    except IOError:
        return(1, 'IOError on input file: ' + input_file)
        
    keys = None
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            text = (cell.text for cell in row.cells)
            
            if i == 0:
                keys = tuple(text)
                continue
            
            if set(headers).issubset(keys):
                row_data = dict(zip(keys, text))
                tables[ansible_table_name].append(row_data)
    
    result['ansible_facts'] = tables

    return(0, result)

def main():
    module_args = dict(
        src = dict(type='str', required=True),
        name = dict(type='str', required=True),
        headers = dict(type=list, required=True)
    )

    result = dict(
        changed=False,
        facts_json='',
        message=''
    )

    module = AnsibleModule(
        argument_spec=module_args,
        supports_check_mode=False
    )
    if not HAS_LIB:
        module.fail_json(msg='Missing required library: python-docx')
        
    ret_code, response = parse_tables_dict(module.params['src'], module.params['name'], module.params['headers'])
    
    if ret_code == 1:
        result['message'] = response
        module.fail_json(**result)
    else:
        result['facts_json'] = response
        result['message'] = 'Done'
    
    module.exit_json(**result)

if __name__ == '__main__':
    main()
