import arrow, os
import json

def folder_cleaner(path):
    criticalTime = arrow.now().shift(hours=+5).shift(days=-90)

    for filename in os.listdir(path):
        itemTime = arrow.get(os.stat(os.path.join(path, filename)).st_mtime)
        if itemTime < criticalTime:
            os.remove(os.path.join(path, filename))


def clean_json():
    with open("./json_templates/project_details.json", "w") as project_json:
        json.dump({}, project_json)

    with open("./json_templates/rsp_grid_details.json", "w") as rsp_grid_json:
        json.dump({}, rsp_grid_json)

    with open("./json_templates/rsp_pc_details.json", "w") as rsp_pc_json:    
        json.dump({}, rsp_pc_json)
