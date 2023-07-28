from pywebio.input import *
from pywebio.output import *
from pywebio.session import *
from pywebio.pin import *
from pywebio import start_server

import urllib.request
import json
 

def read_datatable(instance_id):
    """return the current data in the datatable"""
    return eval_js("""
    window[`ag_grid_${instance_id}_promise`].then(function(gridOptions) {
        function unflatten_row(row) {
            let result = {};
            Object.keys(row).forEach((field) => {
                let path = gridOptions.field2path(field);
                let val = row[field];
                let current = result;
                for (let i = 0; i < path.length - 1; i++) {
                    if (!(path[i] in current)) {
                        current[path[i]] = {};
                    }
                    current = current[path[i]];
                }
                current[path[path.length - 1]] = val;
            });
            return result;
        }
        let rows = []
        gridOptions.api.forEachLeafNode((row)=> rows.push(unflatten_row(row.data)))
        return rows;
    })""", instance_id=instance_id)

with urllib.request.urlopen(
        'https://fakerapi.it/api/v1/persons?_quantity=30') as f:
    data = json.load(f)['data']

put_datatable(data,
                instance_id='user',
              grid_args={
                  'defaultColDef': {
                      'editable': True,
                  },
              })

put_button(
    "read_datatable", 
    lambda: put_code(json.dumps(read_datatable('user'), indent=2, ensure_ascii=False))
)