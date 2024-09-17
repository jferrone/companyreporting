import os
import sqlite3
import argparse
import numpy as np
import pandas as pd
import datetime as dt
import plotly.express as px
import plotly.graph_objects as go

ID_BLUE = '#3a547c'         # Company blue
ID_RED = '#ad2e38'          # Company red
ID_GREY = '#666666'         # Company grey (extra color)

pd.set_option('mode.chained_assignment', None)

def bytes_to_kb(bytes): return (bytes / 1024)           # Bytes to KBs
def bytes_to_mb(bytes): return (bytes / (1024 ** 2))    # Bytes to MBs
def bytes_to_gb(bytes): return (bytes / (1024 ** 3))    # Bytes to GBs
def bytes_to_tb(bytes): return bytes / (1024 ** 4)      # Bytes to TBs

def sql_query(path):
    conn = sqlite3.connect(path) # This is for testing purposes

    query = """
        SELECT FullPath as Path,FileSizeBytes as Size, "File" as Type FROM files
        UNION
        SELECT FullPath as Path,FolderSizeBytes as Size, "Folder" as Type FROM folders
        ORDER BY Type
    """

    files_df = pd.read_sql_query('SELECT * FROM files', conn)                       # Reading in the files table
    folders_df = pd.read_sql_query('SELECT * FROM folders', conn)                   # Reading in the folders table
    summary_df = pd.read_sql_query('SELECT * FROM summary', conn)                   # Reading in the summary table
    details_df = pd.read_sql_query('SELECT * FROM details', conn)                   # Reading in the details table
    dir_tree_df = pd.read_sql_query(query, conn)                                    # DirTree table, combining 'files' and 'folders'
    ritm_num = pd.read_sql_query('SELECT JobID FROM summary', conn)['JobID'][0]     # Grabbing the RITM number

    conn.close()
    
    return dir_tree_df, details_df, ritm_num, files_df, folders_df, summary_df

def generate_graphs(files_df):
    # Cleaning the data
    files_df['FileLastModified'] = pd.to_datetime(files_df['FileLastModified'])
    files_df['FileCreationDate'] = pd.to_datetime(files_df['FileCreationDate'])
    files_df['Year'] = files_df['FileCreationDate'].dt.year
    files_df['YearMonth'] = files_df['FileCreationDate'].dt.strftime('%Y/%m')
    files_df['YearMonthDay'] = files_df['FileCreationDate'].dt.strftime('%Y/%m/%d')
    files_df['FileSizeMB'] = bytes_to_mb(files_df['FileSizeBytes'])
    files_df['FileExtension'] = files_df['FileExtension'].replace('', 'NULL')

    # Extensions Counts DF
    ext_counts_df = files_df['FileExtension'].value_counts().reset_index()
    ext_counts_df.columns = ['Extension', 'Count']

    # Top Ten Files (Legacy FCR)
    files_topten_df = files_df[['FileName', 'FileSizeGB']].sort_values('FileSizeGB', ascending=False).head(10)
    files_topten_df['FileSizeGB'] = np.round(files_topten_df['FileSizeGB'], 4)

    # Dates Counts DFs
    year_counts_df = files_df['Year'].value_counts().reset_index()
    year_counts_df.columns = ['Date', 'Count']
    yearmonth_counts_df = files_df['YearMonth'].value_counts().reset_index()
    yearmonth_counts_df.columns = ['Date', 'Count']
    yearmonthday_counts_df = files_df['YearMonthDay'].value_counts().reset_index()
    yearmonthday_counts_df.columns = ['Date', 'Count']

    # Extensions GBs DF
    ext_gbs_df = files_df[['FileExtension', 'FileSizeGB']].groupby(['FileExtension']).sum().reset_index().sort_values(['FileSizeGB'], ascending=False)

    # Trinary? DF, Greater than 10, Greater than 1MB and Less than 10MB or Less than 1MB 
    def size_groups(FileSizeMB):
        if FileSizeMB <= 1:
            return 'less than or equal to 1MB'
        elif FileSizeMB > 1 and FileSizeMB <= 10:
            return 'less than or equal to 10MB and greater than 1MB'
        elif FileSizeMB > 10:
            return 'greater than 10MB'

    size_df = files_df['FileSizeMB']
    size_df['FileSizeMB'] = size_df.apply(size_groups)
    size_df = size_df['FileSizeMB'].value_counts().reset_index()
    size_df.columns = ['labels', 'counts']

    # Extensions DF for Grouped Table
    ext_df = files_df[['FileExtension', 'FileType', 'FileFormat', 'Class']]
    ext_df.loc[:, 'FileExtensionCounts'] = ext_df.groupby(['FileExtension']).__getitem__('FileFormat').transform('count')
    ext_grouped_df = ext_df.sort_values(
        ['FileExtension', 'FileType', 'FileFormat', 'Class']
    ).groupby(
        ['FileExtension', 'FileExtensionCounts', 'FileType', 'Class', 'FileFormat']
    ).size()

    ext_bar = px.bar(
        ext_counts_df, x = 'Extension', y = 'Count', template = 'ggplot2', text_auto = ''
        ).update_layout(
            font_family = 'Montserrat, sans-serif',
            title_text = '<b>File Count by Extension</b>',
            xaxis_title = '',
            yaxis_title = '',
            xaxis = dict(tickangle = 45),
            yaxis = dict(tickformat = ','),
            margin = dict(b=10, l=10, r=10,t=50),
            title=dict(
                font_color='black',
                font_size=16,
                x=0.95,
                y=0.96,
                xanchor='right'
            )
        ).update_yaxes(
            showgrid = True
        ).update_traces(
            marker_color = ID_BLUE,
            textposition = 'outside',
            cliponaxis = False,
            hovertemplate = '%{y} %{x} files'
        ).to_html(
            full_html=False,
            include_plotlyjs='cdn',
            config={
                'displaylogo':False,
                'modeBarButtonsToRemove': ['toImage', 'lasso2d']
            }
        )

    year_bar = px.bar(
        year_counts_df, x = 'Date', y = 'Count', template = 'ggplot2', text_auto = ''
        ).update_layout(
            font_family = 'Montserrat, sans-serif',
            title_text = '<b>File Count by Year</b>',
            xaxis_title = '',
            yaxis_title = '',
            yaxis = dict(tickformat = ','),
            margin = dict(b=10, l=10, r=10, t=50),
            title=dict(
                font_color='black',
                font_size=20,
                x=0.95,
                y=0.96,
                xanchor='right'
            )
        ).update_xaxes(
            tickformat = "%Y",
            tickangle = 45,
            dtick = 'M12'
        ).update_yaxes(
            showgrid = True
        ).update_traces(
            marker_color = ID_BLUE,
            textposition = 'outside',
            cliponaxis = False,
            hovertemplate = '%{y} files created in %{x}'
        ).to_html(
            full_html=False,
            include_plotlyjs='cdn',
            config={
                'displaylogo':False,
                'modeBarButtonsToRemove': ['toImage', 'lasso2d']
            }
        )

    size_pie = px.pie(
        size_df, values = 'counts', names = 'labels', title = '<b>File Count by Size</b>', template = 'ggplot2'
        ).update_traces(
            hovertemplate = '%{value} files where %{label}',
            marker = dict(
                colors = [ID_BLUE, ID_RED, ID_GREY]
                )
        ).update_layout(
            legend=dict(
                orientation='h',
                xanchor='center',
                yanchor='top',
                x=0.5,
                y=0,
                font_size=12
            ),
            title=dict(
                font_color='black',
                font_size=20,
                x=0.95,
                y=0.96,
                xanchor='right'
            ),
            margin = dict(b=10, l=10, r=10, t=50),
            autosize = True
        ).to_html(
            full_html=False,
            include_plotlyjs='cdn',
            config={
                'displaylogo':False,
                'modeBarButtonsToRemove': ['toImage', 'lasso2d']
            }
        )

    ext_gbs_bar = px.bar(
        ext_gbs_df, x = 'FileExtension', y = 'FileSizeGB', template = 'ggplot2', text_auto = '.4f'
        ).update_layout(
            font_family = 'Montserrat, sans-serif',
            title_text = '<b>Total Size (GB) by Extension</b>',
            xaxis_title = '',
            yaxis_title = '',
            xaxis = dict(tickangle = 45),
            yaxis = dict(tickformat = ','),
            margin = dict(b=10, l=10, r=10, t=50),
            title=dict(
                font_color='black',
                font_size=20,
                x=0.95,
                y=0.96,
                xanchor='right'
            )
        ).update_yaxes(
            showgrid = True
        ).update_traces(
            marker_color = ID_BLUE,
            textposition = 'outside',
            textangle = -30,
            cliponaxis = False,
            hovertemplate = '%{y:.4f} GBs of %{x} files'
        ).to_html(
            full_html=False,
            include_plotlyjs='cdn',
            config={
                'displaylogo':False,
                'modeBarButtonsToRemove': ['toImage', 'lasso2d']
            }
        )

    files_topten_tbl = go.Figure(
        data=go.Table(
            header=dict(values=list(['<b>Top 10 Files</b>', '<b>Size (GB)</b>'])),
            cells=dict(
                values=[files_topten_df['FileName'], files_topten_df['FileSizeGB']], 
                height=25
            )
            )).update_layout(
                title_text = '<b>Top 10 Files by Size</b>',
                margin=dict(b=10, l=10, r=10, t=50),
                autosize=True,
                title=dict(
                    font_color='black',
                    font_size=20,
                    x=0.95,
                    y=0.96,
                    xanchor='right'
                ), template='ggplot2'
        ).to_html(
            full_html=False, 
            include_plotlyjs='cdn',
            config={
                'displaylogo':False,
                'modeBarButtonsToRemove': ['toImage', 'lasso2d']
            }
        )

    html_graphs = ext_bar, year_bar, size_pie, ext_gbs_bar, files_topten_tbl

    return html_graphs


def build_tree_dict(paths, types, sizes_bytes, sizes_kb, sizes_mb, sizes_gb):
    tree = {}
    for path, path_type, size_bytes, size_kb, size_mb, size_gb in zip(paths, types, sizes_bytes, sizes_kb, sizes_mb, sizes_gb):
        if path.startswith("\\\\"):
            path = path.replace("\\\\", "")
        parts = path.split('\\')
        current = tree
        for i, part in enumerate(parts):
            if part not in current:
                # Assign 'file' or 'folder' type based on Type column
                current[part] = {
                    'type': path_type if i == len(parts) - 1 else 'folder', 
                    'children': {},
                    'size_bytes': size_bytes if i == len(parts) - 1 else 0,
                    'size_kb': size_kb if i == len(parts) - 1 else 0,
                    'size_mb': size_mb if i == len(parts) - 1 else 0,
                    'size_gb': size_gb if i == len(parts) - 1 else 0
                }
            current = current[part]['children']
    return tree

def build_html_tree(tree, path=""):
    html = ""
    for name, data in tree.items():
        item_type = data['type']
        subtree = data['children']
        size_bytes = data['size_bytes']
        size_kb = data['size_kb']
        size_mb = data['size_mb']
        size_gb = data['size_gb']

        size_display = f'''
            <span class="size size-bytes">&emsp;<b>{size_bytes} bytes</b></span>
            <span class="size size-kb" style="display: none;">&emsp;<b>{size_kb:.6f} KB</b></span>
            <span class="size size-mb" style="display: none;">&emsp;<b>{size_mb:.6f} MB</b></span>
            <span class="size size-gb" style="display: none;">&emsp;<b>{size_gb:.6f} GB</b></span>
        '''

        if item_type == 'folder':  # It's a folder
            folder_id = f"folder_{path.replace('/', '_')}_{name}"
            if subtree:  # Check if it has items inside
                html += f"""
                <li>
                    <span class="folder" onclick="toggleFolder('{folder_id}')">{name} {size_display}</span>
                    <ul class="nested" id="{folder_id}">
                        {build_html_tree(subtree, f"{path}/{name}")}
                    </ul>
                </li>
                """
        else:  # It's a file
            html += f"""
            <li>
                <span class="file-icon"></span> {name} {size_display}
            </li>
            """
    return html

def propagate_sizes(tree):
    total_size = 0

    # Recursively compute sizes
    for key, item in tree.items():
        if item['type'] == 'folder':
            folder_size = propagate_sizes(item['children'])
            item['size_bytes'] += folder_size
            item['size_kb'] += folder_size / 1024
            item['size_mb'] += folder_size / 1024 ** 2
            item['size_gb'] += folder_size / 1024 ** 3
        total_size += item['size_bytes']

    return total_size

def generate_html_fcr(graph_html, details_df, totals_tbl):
    html_code = f"""
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style type="text/css" media="screen">
            body {{
                font-family: 'Calibri', san-serif;
            }}
            .descriptor {{
                text-align: right;
            }}
            .table {{
                height: 75px;
                overflow: hidden;
            }}
            .floatLeft {{
                display: inline-block;
            }}
            .container {{
                display: flex;
            }}
            .logo_image {{
                position: absolute;
                height: 50%;
                width: auto;
                max-height: 75px;
                right: 50px;
            }}
            .totals{{
                display: flex;
                height: 75px;
            }}
            .plots{{
                display: flex;
                justify-content: center;
                align-items: center;
                flex-wrap: wrap;
                margin: 0 auto;
                margin-top: 10px;
            }}
            .graph1, .graph2, .graph3, .graph4, .responsive-table-container {{
                margin: 0.5px;
                box-sizing: border-box;
            }}
            .graph1, .graph2{{
                height: 50%;
                width: 45%;
                display: inline-block;
            }}
            .graph3, .graph4, .responsive-table-container{{
                display: inline-block;
                height: 50%;
            }}
            .graph3{{
                width: 45%;
            }}
            .graph4, .responsive-table-container{{
                width: 22.46%;
                display: inline-block;
                white-space:no-wrap;
            }}
            .responsive-table-container table {{
                width: 100%;
                table-layout: auto; /* Allow dynamic column widths */
            }}
        </style>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    </head>
    <body>
        <hr style="background-color:#96131d; height:10px;">
        <div class="header">
            <table class="table floatLeft">
                <tbody>
                    <tr>
                        <td class="descriptor">
                            <b>Client Name:</b> 
                        </td>
                        <td>{details_df.loc[0, 'ClientName']}</td>
                        <td class="descriptor">
                            <b>Date:</b> 
                        </td>
                        <td>{details_df.loc[0, 'Date']}</td>
                    </tr>
                    <tr>
                        <td class="descriptor">
                            <b>Matter Name:</b>
                        </td>
                        <td>{details_df.loc[0, 'MatterName']}</td>
                        <td class="descriptor">
                            <b>Evidence ID:</b>
                        </td>
                        <td>{details_df.loc[0, 'EvidenceId']}</td>
                    </tr>
                    <tr>
                        <td class="descriptor">
                            <b>Custodian Name:</b>
                        </td>
                        <td>{details_df.loc[0, 'CustodianName']}</td>
                        <td class="descriptor">
                            <b>Project Manager:</b>
                        </td>
                        <td>{details_df.loc[0, 'ProjectManager']}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        <hr style="background-color:#96131d; height:10px;">
        <div class="totals">
            {totals_tbl}
        </div>
        <div class="plots">
            <div class="graph1">
                {graph_html[0]}
            </div>
            <div class="graph2">
                {graph_html[1]}
            </div>
            <div class="graph3">
                {graph_html[3]}
            </div>
            <div class="graph4">
                {graph_html[2]}
            </div>
            <div class="responsive-table-container">
                {graph_html[4]}
            </div>
        </div>
    </body>
</html>
    """
    return html_code

def generate_html_dirtree(paths, types, sizes_bytes, sizes_kb, sizes_mb, sizes_gb, details_df, totals_tbl):
    tree = build_tree_dict(paths, types, sizes_bytes, sizes_kb, sizes_mb, sizes_gb)
    propagate_sizes(tree)
    html_structure = build_html_tree(tree)

    html_code = f"""
<!DOCTYPE html>
<html>
    <head>
        <style type="text/css" media="screen">
            ul, #dirTree {{
                list-style-type: none;
            }}
            body {{
                font-family: 'Calibri', san-serif;
            }}
            .folder {{
                cursor: pointer;
                user-select: none; /* Prevent text selection */
            }}
            .folder::before {{
                content: "\\1F4C1"; /* Closed Folder */
                display: inline-block;
                margin-right: 5px;
            }}
            .folder-open::before {{
                content: "\\1F4C2"; /* Open Folder */
                display: inline-block;
                margin-right: 5px;
            }}
            .nested {{
                display: none;
            }}
            .active {{
                display: block;
            }}
            .file-icon::before {{
                content: "\\1F4C4";
                display: inline-block;
                margin-right: 5px;
            }}
            .folder-icon::before {{
                content: "\\1F4C1"; /* Folder icon */
                display: inline-block;
                width: 16px;
                height: 16px;
                color: black; /* Color of the icon */
                margin-right: 6px;
                font-size: 16px; /* Size of the icon */
            }}
            .descriptor {{
                text-align: right;
            }}
            .floatLeft {{
                display: inline-block;
            }}
            .table {{
                height: 75px;
                overflow: hidden;
            }}
            .container {{
                display: flex;
            }}
            .logo_image {{
                position: absolute;
                height: 50%;
                width: auto;
                max-height: 75px;
                right: 50px;
            }}
            .totals{{
                display: flex;
                height: 75px;
            }}
        </style>
        <script>
            function toggleFolder(id) {{
            var element = document.getElementById(id);
            var caret = element.previousElementSibling;
            element.classList.toggle("active");
            caret.classList.toggle("folder-open");
            }}
            function toggleSizeFormat(format) {{
            var bytes = document.querySelectorAll('.size-bytes');
            var kbs = document.querySelectorAll('.size-kb');
            var mbs = document.querySelectorAll('.size-mb');
            var gbs = document.querySelectorAll('.size-gb');
            if (format === 'bytes') {{
                bytes.forEach(el => el.style.display = '');
                kbs.forEach(el => el.style.display = 'none');
                mbs.forEach(el => el.style.display = 'none');
                gbs.forEach(el => el.style.display = 'none');
            }} else if (format === 'kbs') {{
                bytes.forEach(el => el.style.display = 'none');
                kbs.forEach(el => el.style.display = '');
                mbs.forEach(el => el.style.display = 'none');
                gbs.forEach(el => el.style.display = 'none');
            }} else if (format === 'mbs') {{
                bytes.forEach(el => el.style.display = 'none');
                kbs.forEach(el => el.style.display = 'none');
                mbs.forEach(el => el.style.display = '');
                gbs.forEach(el => el.style.display = 'none');
            }} else {{
                bytes.forEach(el => el.style.display = 'none');
                kbs.forEach(el => el.style.display = 'none');
                mbs.forEach(el => el.style.display = 'none');
                gbs.forEach(el => el.style.display = '');
            }}
            }}
        </script>
    </head>
    <body>
        <hr style="background-color:#96131d; height:10px;">
        <div class="header">
            <table class="table floatLeft">
                <tbody>
                    <tr>
                        <td class="descriptor">
                            <b>Client Name:</b> 
                        </td>
                        <td>{details_df.loc[0, 'ClientName']}</td>
                        <td class="descriptor">
                            <b>Date:</b> 
                        </td>
                        <td>{details_df.loc[0, 'Date']}</td>
                    </tr>
                    <tr>
                        <td class="descriptor">
                            <b>Matter Name:</b>
                        </td>
                        <td>{details_df.loc[0, 'MatterName']}</td>
                        <td class="descriptor">
                            <b>Evidence ID:</b>
                        </td>
                        <td>{details_df.loc[0, 'EvidenceId']}</td>
                    </tr>
                    <tr>
                        <td class="descriptor">
                            <b>Custodian Name:</b>
                        </td>
                        <td>{details_df.loc[0, 'CustodianName']}</td>
                        <td class="descriptor">
                            <b>Project Manager:</b>
                        </td>
                        <td>{details_df.loc[0, 'ProjectManager']}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        <hr style="background-color:#96131d; height:10px;">
        <div class="totals">
            {totals_tbl}
        </div>
        <div>
            <h2>File Directory Navigator</h2>
            <h3>Size Type
                <select id="sizeFormatDropdown" onchange="toggleSizeFormat(this.value)">
                    <option value="bytes">Bytes</option>
                    <option value="kbs">KB</option>
                    <option value="mbs">MB</option>
                    <option value="gbs">GB</option>
                </select>
            </h3>
            <ul id="dirTree">
                {html_structure}
            </ul>
        </div>
    </body>
</html>
"""
    return html_code

def main():
    # Initializing the Argument Parser
    parser = argparse.ArgumentParser(description='Generates a First Contact and Directory Tree Report')

    # Add command-line arguments
    parser.add_argument('-db', '--database', type=str, required=True, help='Path to .db file')
    parser.add_argument('-op', '--outpath', type=str, required=False, default=f'{os.environ["ProgramData"]}', help='Output path')
    parser.add_argument('-nodt', '--nodirectorytree', action='store_true', help='Exclude the Directory Tree Report')
    parser.set_defaults(exclude_dt = False)

    # Parse the arguments
    args = parser.parse_args()
    db_file = args.database
    output = args.outpath
    exclude_dt = args.nodirectorytree

    # SQL Query
    dir_tree_df, details_df, ritm_num, files_df, folders_df, summary_df = sql_query(db_file)

    # For the file name.
    evidence_num = details_df.loc[0, 'EvidenceId']

    # Extract paths and types from the DataFrame
    paths = dir_tree_df['Path'].tolist()
    types = dir_tree_df['Type'].tolist()
    sizes_bytes = dir_tree_df['Size'].tolist()
    sizes_kb = bytes_to_kb(dir_tree_df['Size']).tolist()
    sizes_mb = bytes_to_mb(dir_tree_df['Size']).tolist()
    sizes_gb = bytes_to_gb(dir_tree_df['Size']).tolist()

    date = dt.datetime.now().strftime('%Y%m%d-%H%M%S')

    totalfiles = summary_df.loc[0, 'TotalFiles']
    totalgbs = np.round(summary_df.loc[0, 'TotalSizeGB'], 4)

    totals_tbl = go.Figure(data=[go.Table(
        header=dict(values=['<b>Total Files</b>', '<b>Size (GB)</b>']),
        cells = dict(  
            values=[totalfiles, totalgbs], height=25
            ))]
        ).update_layout(
            margin=dict(b=0, l=10, r=10,t=10), template='ggplot2',
            width=300
    ).to_html(
            full_html=False,
            include_plotlyjs='cdn',
            config={
                'displaylogo':False,
                'modeBarButtonsToRemove': ['toImage', 'lasso2d']
            }
    )

    output_ritm = os.path.join(output, 'Generic', 'Reports', f'{ritm_num}', f'{evidence_num}')
    os.makedirs(output_ritm, exist_ok=True)
    fcrname = f'{ritm_num}_{evidence_num}_FirstContactReport_{date}.html'
    dirtreename = f'{ritm_num}_{evidence_num}_DirTreeReport_{date}.html'
    fcr_html_path = os.path.join(output_ritm, fcrname)

    print('\nGenerating First Contact Report')
    # Generate the HTML code for the First Contact Report
    graph_html = generate_graphs(files_df)
    fcreport_html_output = generate_html_fcr(graph_html, details_df, totals_tbl)
    print('Writing First Contact Report to HTML file.')

    with open(fcr_html_path, "w") as f:
        f.write(fcreport_html_output)

    print(f"\nFirst Contact Report generated as: \n{fcrname} \n\nReports generated here: \n{fcr_html_path}\n")

    if exclude_dt is False:
        print('\nGenerating Directory Tree')
        # Generate the HTML code
        dirtree_html_output = generate_html_dirtree(paths, types, sizes_bytes, sizes_kb, sizes_mb, sizes_gb, details_df, totals_tbl)

        print('Writing Directory Tree Report to HTML file.')
        # Save the HTML code to a file

        dirtree_html_path = os.path.join(output_ritm, dirtreename)

        with open(dirtree_html_path, "w") as f:
            f.write(dirtree_html_output)

        print(f"\nDirectory Tree generated as: \n{dirtreename} \n\nReports generated here: \n{dirtree_html_path}")

if __name__ == "__main__": main()