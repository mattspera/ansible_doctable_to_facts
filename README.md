# doctable_to_facts

Read and parse specific tables within a .docx file and output Ansible facts.

To execute example and view Ansible facts output:

`ansible-playbook parse-table.yml -vvv`

For use within playbook, place `doctable_to_facts.py` into the **library/** directory within your Ansible project directory.
