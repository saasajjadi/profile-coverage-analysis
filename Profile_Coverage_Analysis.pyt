# -*- coding: utf-8 -*-
"""Profile coverage estimate by profile name or nearest line, v6.3.
This Python toolbox is provided as a starting point; test in ArcGIS Pro with
copies of actual data. A span between observations is not proof that all of
the intervening route was traversed or geologically inspected.
"""
TERMS_OF_USE = "TERMS OF USE — Profile Coverage Analysis\n\nDeveloper: سید علی اکبر سجادی\nEmail: SaSajjadi@outlook.com\nLinkedIn: https://www.linkedin.com/in/aliakbar-sajjadi/\nGitHub: https://github.com/saasajjadi\n\n1. Personal, educational, research, and commercial use of this tool is permitted.\n2. Modification of the source code and redistribution of either the original or modified version are permitted, provided that the original developer's name is retained. Changes must be identified, and a modified version must not be presented as the official version.\n3. Results are estimates and must be checked against field data and professional judgment. Covered does not, by itself, establish continuous traversal or complete geological coverage.\n4. The tool is provided without a warranty that it is suitable for every project. Users are responsible for checking inputs, settings, and results before using them for decisions.\n5. ArcGIS Pro and any other required software are subject to their own terms and licenses."
PARAMETER_GUIDE = 'PROFILE COVERAGE ANALYSIS — SHORT PARAMETER GUIDE\nDeveloper: سید علی اکبر سجادی | Version 6.3\n\nINPUTS\n1. Observation points: Select the point layer for the reporting period. Active selection and definition query apply; dates are not filtered automatically.\n2. Points - profile name field (name mode): Select the text field containing the profile name for each point. Required only for By Profile Name.\n3. Profile lines: Select one line layer, with one valid, singlepart, nonzero-length feature per profile. Active selection and definition query apply.\n4. Profiles - profile name field (name mode): Select the text field identifying each profile; names must be unique and nonblank. Required only for By Profile Name.\n5. Optional profile grouping field: Select a text field on the profile layer, such as material, to produce group summaries. Leave blank for the overall total only.\n6. Output file geodatabase: Select an existing writable file geodatabase ending in .gdb, not an ordinary folder.\n7. Folder for Excel and run notes: Select an existing writable folder for the XLSX report and TXT run notes.\n8. Point assignment method: By Profile Name matches the two text fields. By Nearest Line assigns points to the nearest selected line and creates PROFILE_<source OID> labels in outputs.\n9. Optional max distance in metres: Used only in nearest-line mode. Enter a positive distance to reject points farther away, or leave blank for no distance limit. Inspect distances and ties.\n10. Optional exact map name: Enter an exact map name to add covered lines to that map; otherwise leave blank. Review symbology and save the project manually.\n11. Optional exact layout name: Enter the exact layout name only if updating a layout text element. Provide the text element name as well.\n12. Optional exact layout text element name: Enter the exact name of an existing text element in the specified layout. The tool replaces its text. Review layout and save the project manually.\n13. Analysis coordinate system: Choose an appropriate projected CRS with metre units. If blank, the tool uses the output CRS environment, otherwise the line CRS if projected/metre, then the point CRS if projected/metre. Verify suitability for the study area.\n14. Optional point geographic transformation: If required, enter the exact ArcGIS transformation name for point-source CRS to analysis CRS; leave blank when none is required. Select a method appropriate for the data and area.\n15. Optional profile geographic transformation: Same as above, for the profile line-source CRS. A missing or invalid required transformation stops the tool and reports candidates.\n\nOUTPUTS (GENERATED AUTOMATICALLY)\nProfile statistics: Per-profile line, source OID, total and estimated covered/uncovered lengths and percentages, observation count and QA distances.\nPoint audit: Valid-geometry point output with assignment status, source/profile IDs, offsets and nearest-line diagnostics. Points with missing geometry are in Point issues instead.\nCovered lines: Estimated line segments between the first and last matched point positions.\nSummary table: Group totals if a grouping field is selected, plus the overall total.\nPoint issues: Source OIDs and status for missing geometry, missing/unmatched names, nearest ties, and points outside a specified search radius.\nExcel report: Formatted coverage summary and notes; editing its cells does not change GIS outputs.\nRun notes: Input, CRS, transformation, assignment and limitations for the run.\n\nINTERPRETATION AND CAUTIONS\nCovered is the estimated distance along a profile between its first and last matched observations; gaps between observations are included. Uncovered is the remainder, not confirmed untraversed distance. Zero or one matched point gives zero estimated span. Overlapping profiles are counted separately. Results are 2D planar; long projected lines are not automatically densified. Review CRS, assignments and source geometry. Automatically generated OID labels may change if source OIDs change.\n\nFor on-screen help, open the separate "Parameter Guide | راهنمای فیلدهای ابزار" tool in this toolbox and click Run. The guide appears in its Geoprocessing messages. This is not a parameter tooltip.'

import os
import math
import html
import unicodedata
import uuid
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import arcpy

VERSION = '6.3'
AUTHOR = 'سید علی اکبر سجادی'
EMAIL = 'SaSajjadi@outlook.com'
LINKEDIN = 'https://www.linkedin.com/in/aliakbar-sajjadi/'
GITHUB = 'https://github.com/saasajjadi'
BY_NAME = 'By Profile Name'
BY_NEAREST = 'By Nearest Line'
TIE_TOLERANCE_M = 1e-8


def clean(value):
    return '' if value is None else str(value).strip()


def profile_key(value):
    return unicodedata.normalize('NFKC', clean(value)).casefold()


def field_info(layer, name, required=True):
    if not name:
        if required:
            raise ValueError('A profile-name field is required for {}.'.format(layer))
        return None
    matches = [f for f in arcpy.ListFields(layer) if f.name.casefold() == name.casefold()]
    if len(matches) != 1:
        raise ValueError('Field {} is missing or ambiguous in {}. Select the actual field name.'.format(name, layer))
    if matches[0].type != 'String':
        raise ValueError('Field {} in {} must be a text field.'.format(name, layer))
    return matches[0].name


def write_excel_report(path, groups, grouped, total_m, covered_m,
                       profile_count, point_count, run_id, crs_name,
                       point_layer, point_field, line_layer, line_field, group_field,
                       method, max_distance_m, crs_details=None):
    """Create a readable, formula-backed report with initial cached values.

    Excel recalculates formulas after manual edits; those edits do not change
    the GIS or prove continuous traversal. Geodatabase outputs remain the
    authoritative snapshot of the geoprocessing run.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.workbook.properties import CalcProperties
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    wb.properties.creator = AUTHOR
    wb.properties.title = 'Profile Coverage Analysis'
    wb.calculation = CalcProperties(calcMode='auto', fullCalcOnLoad=True, forceFullCalc=True)
    ws = wb.active
    ws.title = 'Coverage'
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'B5'
    ws.merge_cells('A1:I1')
    ws['A1'] = 'PROFILE COVERAGE ANALYSIS'
    ws['A1'].font = Font(name='Arial', size=15, bold=True, color='FFFFFF')
    ws['A1'].fill = PatternFill('solid', fgColor='173B58')
    ws['A1'].alignment = Alignment(vertical='center')
    ws.row_dimensions[1].height = 30
    ws.merge_cells('A2:I2')
    ws['A2'] = 'Estimated distance along profiles between first and last matched observations'
    ws['A2'].font = Font(name='Arial', size=10, italic=True, color='526A7A')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[2].height = 27
    headers = ['Group', 'Total (km)', 'Covered (km)', 'Covered (%)',
               'Uncovered (km)', 'Uncovered (%)', 'Profiles', 'Points', 'Check']
    for col, label in enumerate(headers, 1):
        cell = ws.cell(4, col, label)
        cell.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='245577')
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[4].height = 32
    for col, width in zip('ABCDEFGHI', [32, 15, 16, 16, 18, 18, 12, 12, 16]):
        ws.column_dimensions[col].width = width
    cache = {}
    group_rows = []
    thin = Side(style='hair', color='CBDCE5')

    def text_cell(row, col, value):
        c = ws.cell(row, col, value)
        # Force arbitrary source group values to literal strings, not formulas.
        c.data_type = 's'
        return c

    def formula(address, expr, value):
        ws[address] = expr
        cache[address] = value

    def fill_row(row, name, total_km, covered_km, nprofiles, npoints, is_total=False):
        label = name if name else '[Empty group]'
        text_cell(row, 1, 'Total' if is_total else label)
        ws.cell(row, 2, total_km)
        ws.cell(row, 3, covered_km)
        ws.cell(row, 7, nprofiles)
        ws.cell(row, 8, npoints)
        valid = total_km > 0 and 0 <= covered_km <= total_km
        formula('I{}'.format(row),
                '=IF(OR(NOT(ISNUMBER(B{0})),NOT(ISNUMBER(C{0})),B{0}<=0,C{0}<0,C{0}>B{0}),"CHECK","OK")'.format(row),
                'OK' if valid else 'CHECK')
        uncovered = total_km - covered_km
        formula('D{}'.format(row), '=IF(I{0}="OK",C{0}/B{0}*100,"")'.format(row),
                100*covered_km/total_km if valid else '')
        formula('E{}'.format(row), '=IF(I{0}="OK",B{0}-C{0},"")'.format(row),
                uncovered if valid else '')
        formula('F{}'.format(row), '=IF(I{0}="OK",E{0}/B{0}*100,"")'.format(row),
                100*uncovered/total_km if valid else '')
        for col in range(1, 10):
            c = ws.cell(row, col)
            c.font = Font(name='Arial', size=10, bold=is_total,
                          color='13547A' if col in (2, 3) and not is_total else '253A49')
            c.fill = PatternFill('solid', fgColor='DFEDF5' if is_total else ('F3F8FA' if row % 2 else 'FFFFFF'))
            c.border = Border(bottom=thin)
            c.alignment = Alignment(vertical='center', wrap_text=(col == 1),
                                    horizontal='left' if col == 1 else ('center' if col == 9 else 'right'))
            if col in (2, 3, 5):
                c.number_format = '0.00'
            elif col in (4, 6):
                # Values are stored as 0..100; quoted percent is a literal.
                c.number_format = '0.0"%"'
            elif col in (7, 8):
                c.number_format = '0'
        ws.row_dimensions[row].height = max(24, min(180, 15 * (1 + len(label)//28)))

    row = 5
    for group, s in sorted(groups.items()):
        fill_row(row, group, s['total']/1000, s['span']/1000, s['nprof'], s['npoints'])
        group_rows.append(row)
        row += 1
    total_row = row
    fill_row(total_row, 'Total', total_m/1000, covered_m/1000, profile_count, point_count, True)
    if group_rows:
        first, last = group_rows[0], group_rows[-1]
        # Total lengths are blank when any group inputs fail the Check.
        for col, cached in [('B', total_m/1000), ('C', covered_m/1000)]:
            formula('{}{}'.format(col, total_row),
                    '=IF(COUNTIF(I{0}:I{1},"<>OK")>0,"",SUM({2}{0}:{2}{1}))'.format(first,last,col),
                    cached)
        for col, cached in [('G', profile_count), ('H', point_count)]:
            formula('{}{}'.format(col, total_row), '=SUM({0}{1}:{0}{2})'.format(col, first, last), cached)

    input_rows = group_rows if grouped else [total_row]
    if input_rows:
        # A normal typed entry is checked; the Check formula also catches invalid pasted values.
        dv = DataValidation(type='custom',
                            formula1='=AND(ISNUMBER($B{0}),ISNUMBER($C{0}),$B{0}>0,$C{0}>=0,$C{0}<=$B{0})'.format(input_rows[0]),
                            allow_blank=False)
        dv.errorTitle = 'Invalid length'
        dv.error = 'Total must be positive and Covered must be between zero and Total.'
        dv.showErrorMessage = True
        dv.errorStyle = 'stop'
        ws.add_data_validation(dv)
        dv.add('B{0}:C{1}'.format(input_rows[0], input_rows[-1]))
    footer = total_row + 2
    ws.merge_cells(start_row=footer, start_column=1, end_row=footer, end_column=9)
    ws.cell(footer, 1, 'Estimated results only. See the Notes sheet for definitions and limitations.')
    ws.cell(footer, 1).font = Font(name='Arial', size=9, italic=True, color='526A7A')
    ws.cell(footer, 1).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[footer].height = 28
    ws.merge_cells(start_row=footer+1, start_column=1, end_row=footer+1, end_column=9)
    text_cell(footer+1, 1, 'Prepared by: ' + AUTHOR)
    ws.cell(footer+1, 1).font = Font(name='Arial', size=9, color='526A7A')

    notes = wb.create_sheet('Notes')
    notes.sheet_view.showGridLines = False
    notes.column_dimensions['A'].width = 27
    notes.column_dimensions['B'].width = 98
    note_rows = [
        ('Report', 'Profile Coverage Analysis'),
        ('Prepared by', AUTHOR),
        ('Tool version', VERSION),
        ('Email', EMAIL),
        ('LinkedIn', LINKEDIN),
        ('GitHub', GITHUB),
        ('Terms of use', 'Personal, educational, research and commercial use permitted. Modification and redistribution allowed with original attribution; identify changes and do not present modified versions as official. See the toolbox source or TERMS_OF_USE.txt.'),
        ('Run ID', run_id),
        ('Coordinate system', crs_name),
        ('Point layer / field', point_layer + ' / ' + (point_field or '(not used)')),
        ('Profile layer / field', line_layer + ' / ' + (line_field or '(not used)')),
        ('Optional group field', group_field or '(none)'),
        ('Assignment method', method),
        ('Covered', 'Estimated along-line distance between first and last matched points. Gaps between them remain included.'),
        ('Uncovered', 'Total length minus Covered. This is outside the estimated span, not proven untraversed.'),
        ('Distance rule', ('No maximum offset is applied; even distant points may be included.'
                           if max_distance_m is None else
                           'Nearest-line search radius: {} metres.'.format(max_distance_m))),
        ('Nearest-line QA', 'When using nearest line, inspect Point Audit. Equal nearest distances within 0.00000001 m remain unassigned. Nearly equal distances are recorded, not automatically rejected.'),
        ('Generated IDs', 'Nearest-line mode creates PROFILE_<source OID> labels in outputs. IDs may change if source OIDs change.'),
        ('No or one point', 'A profile with fewer than two matched points has zero estimated Covered; this does not prove it was not traversed.'),
        ('Overlap', 'Overlapping profiles are counted separately. All lengths are 2D planar km.'),
        ('Scope', 'Selections and definition queries apply. Dates are not filtered. The user must select the intended reporting period.'),
        ('Excel edits', 'Changing blue input cells recalculates formulas in spreadsheet software. It does not update GIS layers, project, or map. Re-run the tool for authoritative GIS results.'),
        ('Check column', 'CHECK means Total or Covered is not numeric, Total is not positive, or Covered is outside 0..Total. Fix inputs.'),
        ('Rounding', 'Lengths show 2 decimals and percentages 1 decimal. Values/formulas use full precision; displayed parts may not exactly sum due to rounding.'),
        ('Data quality', 'Review unmatched points, point geometry issues, self-intersections and selected CRS. Do not infer geological coverage or reserves.'),
    ]
    note_rows.extend(crs_details or [])
    for r, (label, value) in enumerate(note_rows, 1):
        label_cell = notes.cell(r, 1, label)
        label_cell.font = Font(name='Arial', size=10, bold=True, color='1F4E68')
        label_cell.alignment = Alignment(vertical='top', wrap_text=True)
        c = notes.cell(r, 2, value); c.data_type = 's'
        c.font = Font(name='Arial', size=10, color='253A49')
        if label in ('Email', 'LinkedIn', 'GitHub'):
            c.hyperlink = ('mailto:' + value) if label == 'Email' else value
            c.font = Font(name='Arial', size=10, color='0563C1', underline='single')
        c.alignment = Alignment(wrap_text=True, vertical='top')
        notes.row_dimensions[r].height = max(22, 15 * (1 + len(value)//88))
    tmp = path + '.building.xlsx'
    cached = path + '.cached.xlsx'
    try:
        wb.save(tmp)
        # openpyxl does not cache formula values. Embed the independently
        # calculated initial values, then request full recalculation on open.
        ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        ET.register_namespace('', ns['x'])
        ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        seen = set()
        with zipfile.ZipFile(tmp, 'r') as zin, zipfile.ZipFile(cached, 'w') as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == 'xl/worksheets/sheet1.xml':
                    root = ET.fromstring(data)
                    for cell in root.findall('.//x:c', ns):
                        address = cell.get('r')
                        if address in cache:
                            if cell.find('x:f', ns) is None:
                                raise ValueError('Missing formula for cached cell {}'.format(address))
                            value = cache[address]
                            node = cell.find('x:v', ns)
                            if node is None:
                                node = ET.SubElement(cell, '{{{}}}v'.format(ns['x']))
                            if isinstance(value, str):
                                cell.set('t', 'str')
                                node.text = value
                            else:
                                cell.attrib.pop('t', None)
                                node.text = str(value)
                            seen.add(address)
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                zout.writestr(info, data)
        if seen != set(cache):
            raise ValueError('Could not set all Excel formula caches.')
        os.replace(cached, path)
    finally:
        for candidate in (tmp, cached):
            if os.path.exists(candidate):
                try: os.remove(candidate)
                except OSError: pass


def build_params():
    def par(label, name, datatype, required='Required', direction='Input'):
        return arcpy.Parameter(displayName=label, name=name, datatype=datatype,
                               parameterType=required, direction=direction)
    points = par('Observation points', 'points', 'GPFeatureLayer')
    points.filter.list = ['Point']
    point_name = par('Points - profile name field (name mode)', 'point_profile_name', 'Field', 'Optional')
    point_name.parameterDependencies = ['points']; point_name.filter.list = ['Text']
    lines = par('Profile lines', 'profiles', 'GPFeatureLayer')
    lines.filter.list = ['Polyline']
    line_name = par('Profiles - profile name field (name mode)', 'line_profile_name', 'Field', 'Optional')
    line_name.parameterDependencies = ['profiles']; line_name.filter.list = ['Text']
    group = par('Optional: profile grouping field (for example, material)', 'group_field', 'Field', 'Optional')
    group.parameterDependencies = ['profiles']; group.filter.list = ['Text']
    method = par('Point assignment method', 'assignment_method', 'GPString')
    method.filter.list = [BY_NAME, BY_NEAREST]
    method.value = BY_NAME
    max_distance = par('Optional: max distance in metres (nearest-line mode)', 'max_distance_m', 'GPDouble', 'Optional')
    ps = [points, point_name, lines, line_name, group,
          par('Output file geodatabase', 'out_gdb', 'DEWorkspace'),
          par('Folder for Excel and run notes', 'out_folder', 'DEFolder'),
          method, max_distance,
          par('Optional: exact map name for span layer', 'map_name', 'GPString', 'Optional'),
          par('Optional: exact layout name', 'layout_name', 'GPString', 'Optional'),
          par('Optional: exact layout text element name', 'text_element', 'GPString', 'Optional'),
          par('Analysis coordinate system (projected, metres)', 'analysis_crs', 'GPCoordinateSystem', 'Optional'),
          par('Optional: point geographic transformation', 'point_transform', 'GPString', 'Optional'),
          par('Optional: profile geographic transformation', 'line_transform', 'GPString', 'Optional')]
    for label, name, datatype in [
        ('Profile statistics', 'profile_stats', 'DEFeatureClass'),
        ('Point audit', 'point_audit', 'DEFeatureClass'),
        ('Covered lines', 'span_lines', 'DEFeatureClass'),
        ('Summary table', 'summary_table', 'DETable'),
        ('Point issues', 'point_issues', 'DETable'),
        ('Excel report', 'excel_report', 'DEFile'),
        ('Run notes', 'run_notes', 'DEFile')]:
        ps.append(par(label, name, datatype, 'Derived', 'Output'))
    return ps


class Toolbox:
    def __init__(self):
        self.label = 'Profile Coverage Analysis | تحلیل پوشش پروفیل‌های پیمایش'
        self.alias = 'generalprofilespan'
        self.tools = [ProfileSpan, ParameterGuide]


class ProfileSpan:
    def __init__(self):
        self.label = 'Profile Coverage Analysis | تحلیل پوشش پروفیل‌های پیمایش'
        self.description = ('Distance along each line from the first to last matched observation; '
                            'intervening gaps are included. Assign by name or nearest line. '
                            'No distance cutoff is applied unless specified in nearest mode. Prepared by: ' + AUTHOR +
                            '. Version: ' + VERSION + '. Email: ' + EMAIL +
                            '. LinkedIn: ' + LINKEDIN + '. GitHub: ' + GITHUB)
        self.canRunInBackground = False

    def getParameterInfo(self):
        return build_params()

    def updateParameters(self, parameters):
        name_mode = parameters[7].valueAsText != BY_NEAREST
        parameters[1].enabled = name_mode
        parameters[3].enabled = name_mode
        parameters[8].enabled = not name_mode
        return

    def updateMessages(self, parameters):
        if parameters[7].valueAsText == BY_NAME:
            for i in (1, 3):
                if not parameters[i].valueAsText:
                    parameters[i].setErrorMessage('Select the name field for name-based assignment.')
        if parameters[7].valueAsText == BY_NEAREST and parameters[8].valueAsText:
            try:
                if float(parameters[8].valueAsText) <= 0:
                    parameters[8].setErrorMessage('The optional maximum distance must be greater than zero.')
            except ValueError:
                parameters[8].setErrorMessage('Enter a numeric distance in metres.')
        return

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
        try:
            # Avoid inherited processing extent or output CRS silently changing the results.
            env_crs = arcpy.env.outputCoordinateSystem
            with arcpy.EnvManager(extent=None, outputCoordinateSystem=None, geographicTransformations=None):
                self._execute(parameters, env_crs)
        except arcpy.ExecuteError:
            arcpy.AddError(arcpy.GetMessages(2))
            arcpy.AddError('A failed run may have left partial outputs. Inspect them before removal.')
            raise
        except Exception as e:
            arcpy.AddError('Analysis stopped: {}'.format(e))
            arcpy.AddError('A failed run may have left partial outputs. Inspect them before removal.')
            raise

    def _execute(self, p, env_crs=None):
        try:
            import openpyxl  # required for formatted Excel report
        except ImportError as exc:
            raise RuntimeError('openpyxl is required in the active ArcGIS Pro Python environment. Install it in a cloned environment using Package Manager.') from exc
        pt, pt_name = p[0].valueAsText, p[1].valueAsText
        line, line_name, group_name = p[2].valueAsText, p[3].valueAsText, p[4].valueAsText
        gdb, folder = p[5].valueAsText, p[6].valueAsText
        method = p[7].valueAsText
        if method not in (BY_NAME, BY_NEAREST):
            raise ValueError('Choose By Profile Name or By Nearest Line.')
        radius_text = p[8].valueAsText if method == BY_NEAREST else None
        max_distance = float(radius_text) if radius_text else None
        if max_distance is not None and (not math.isfinite(max_distance) or max_distance <= 0):
            raise ValueError('Maximum distance must be a positive finite number in metres.')
        map_name, layout_name, text_name = [clean(p[i].valueAsText) for i in (9, 10, 11)]
        if not pt or not line:
            raise ValueError('Choose both the point and profile line layers.')
        if not gdb or not arcpy.Exists(gdb) or not gdb.lower().endswith('.gdb'):
            raise ValueError('Select an existing writable file geodatabase (.gdb).')
        if not folder or not os.path.isdir(folder):
            raise ValueError('Select an existing writable output folder.')
        if bool(layout_name) != bool(text_name):
            raise ValueError('Enter both layout name and text element name, or neither.')
        pt_name = field_info(pt, pt_name) if method == BY_NAME else None
        line_name = field_info(line, line_name) if method == BY_NAME else None
        group_name = field_info(line, group_name, required=False)
        pdesc, ldesc = arcpy.Describe(pt), arcpy.Describe(line)
        if pdesc.shapeType != 'Point' or ldesc.shapeType != 'Polyline':
            raise ValueError('Observations must be points and profiles must be polylines.')
        point_sr, line_sr = pdesc.spatialReference, ldesc.spatialReference
        for label, source in [('Points', point_sr), ('Profiles', line_sr)]:
            if not source or source.name == 'Unknown' or source.type not in ('Projected', 'Geographic'):
                raise ValueError(label + ' have an unknown coordinate system. Identify the actual source CRS and define it correctly before running. Do not guess.')

        def as_sr(value, text=None):
            if hasattr(value, 'type') and hasattr(value, 'exportToString'):
                return value
            text = text or (str(value) if value else '')
            if not text:
                return None
            if text.isdigit():
                return arcpy.SpatialReference(int(text))
            if text.lstrip().startswith(('PROJCS[', 'GEOGCS[', 'PROJCRS[', 'GEOGCRS[', 'GEODCRS[')):
                obj = arcpy.SpatialReference()
                obj.loadFromString(text)
                return obj
            return arcpy.SpatialReference(text)

        selected_sr = as_sr(p[12].value, p[12].valueAsText) if p[12].valueAsText else None
        inherited_sr = as_sr(env_crs) if env_crs else None
        sr = selected_sr or inherited_sr or (line_sr if line_sr.type == 'Projected' and abs(line_sr.metersPerUnit-1.0) < 1e-6 else None) or (point_sr if point_sr.type == 'Projected' and abs(point_sr.metersPerUnit-1.0) < 1e-6 else None)
        if not sr or sr.type != 'Projected' or abs(sr.metersPerUnit-1.0) > 1e-6:
            raise ValueError('Choose an analysis coordinate system that is projected in metres and suitable for your study area. Do not use a geographic CRS or a projected CRS in feet.')
        arcpy.AddMessage('Analysis CRS: {}. Check its area of use and suitability for the site.'.format(sr.name))

        def transform_for(label, source, target, requested, extent):
            # Compare datums, not full CRS objects: XY precision and coordinate system
            # descriptions can differ without requiring a datum transformation.
            source_gcs = source.GCS if source.type == 'Projected' else source
            target_gcs = target.GCS if target.type == 'Projected' else target
            sc = getattr(source_gcs, 'datumCode', 0) or 0
            tc = getattr(target_gcs, 'datumCode', 0) or 0
            same_datum = source_gcs == target_gcs or (sc > 0 and tc > 0 and sc == tc)
            choices = arcpy.ListTransformations(source, target, extent) or []
            if requested:
                if requested not in choices:
                    raise ValueError('{} transformation is not listed as valid for these inputs and extent. Candidates: {}. Check area of use and installed coordinate-system data.'.format(label, ', '.join(choices[:12]) or '(none)'))
                return requested
            if not same_datum:
                raise ValueError('{} may need a geographic datum transformation. Choose a suitable transformation in the {} transformation parameter. Candidates for the input extent: {}. Consult your GIS specialist if uncertain; no transformation is guessed.'.format(label, label.lower(), ', '.join(choices[:12]) or '(none available; check CRS and installed coordinate-system data)'))
            return ''

        pt_transform = transform_for('Point', point_sr, sr, clean(p[13].valueAsText), getattr(pdesc, 'extent', None))
        line_transform = transform_for('Profile', line_sr, sr, clean(p[14].valueAsText), getattr(ldesc, 'extent', None))
        arcpy.AddMessage('Point CRS: {} | geographic transformation: {}'.format(point_sr.name, pt_transform or '(none)'))
        arcpy.AddMessage('Profile CRS: {} | geographic transformation: {}'.format(line_sr.name, line_transform or '(none)'))
        arcpy.AddWarning('Lengths and offsets are 2D planar. Confirm CRS suitability, datum transformations, and input metadata. Projecting line geometries does not densify long segments; use a suitable detailed line if necessary.')
        def projected(geom, source, transformation, label, oid):
            if geom is None or geom.pointCount == 0:
                return geom
            try:
                result = geom if source == sr else geom.projectAs(sr, transformation) if transformation else geom.projectAs(sr)
            except Exception as exc:
                raise ValueError('Could not project {} OID {} to analysis CRS: {}. Verify CRS, area and transformation.'.format(label, oid, exc)) from exc
            if result is None or result.pointCount == 0:
                raise ValueError('Projection produced empty {} OID {}; verify projection area and input coordinates.'.format(label, oid))
            return result

        stamp = datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:8]
        def dest(label):
            return os.path.join(gdb, arcpy.ValidateTableName(label + '_' + stamp, gdb))
        paths = [dest('Profile_Stats'), dest('Point_Audit'), dest('Profile_Covered'),
                 dest('Span_Summary'), dest('Point_Issues')]
        excel = os.path.join(folder, 'Profile_Coverage_' + stamp + '.xlsx')
        notes_path = os.path.join(folder, 'Profile_Coverage_Notes_' + stamp + '.txt')
        if any(arcpy.Exists(v) for v in paths) or any(os.path.exists(v) for v in [excel, notes_path]):
            raise ValueError('An output name already exists; rerun to generate a new run ID.')

        profiles = {}
        profiles_by_oid = {}
        duplicate_names = []
        arcpy.AddMessage('Reading input profiles. Active selections and definition queries apply.')
        line_fields = ['OID@', 'SHAPE@'] + ([line_name] if line_name else []) + ([group_name] if group_name else [])
        with arcpy.da.SearchCursor(line, line_fields) as cursor:
            for row in cursor:
                oid, geom = row[:2]
                geom = projected(geom, line_sr, line_transform, 'profile', oid)
                source_id = str(oid)
                if source_id in profiles_by_oid:
                    raise ValueError('Duplicate profile OID {}; do not use a joined layer with duplicate rows.'.format(oid))
                raw_name = row[2] if line_name else None
                name = clean(raw_name) if method == BY_NAME else 'PROFILE_' + source_id
                k = profile_key(name)
                if method == BY_NAME and not k:
                    raise ValueError('Profile OID {} has no profile name.'.format(oid))
                if len(name) > 255:
                    raise ValueError('Profile name exceeds 255 characters at OID {}.'.format(oid))
                if k in profiles:
                    duplicate_names.append(name)
                    continue
                if geom is None or geom.pointCount < 2 or not math.isfinite(geom.length) or geom.length <= 0:
                    raise ValueError('Profile OID {} has empty, non-finite or zero-length geometry.'.format(oid))
                if geom.partCount != 1:
                    raise ValueError('Profile OID {} is multipart; use one singlepart line per profile.'.format(oid))
                group = clean(row[-1]) if group_name else ''
                if len(group) > 255:
                    raise ValueError('Group value exceeds 255 characters at profile OID {}.'.format(oid))
                profiles[k] = {'name': name, 'oid': source_id, 'geom': geom, 'length': geom.length,
                               'group': group, 'obs': []}
                profiles_by_oid[source_id] = profiles[k]
        if duplicate_names:
            raise ValueError('Each profile must have a unique name after trim/case/Unicode normalization. Duplicate examples: ' + ', '.join(duplicate_names[:10]))
        if not profiles:
            raise ValueError('No selected profile lines were found.')

        point_fields = ['OID@', 'SHAPE@'] + ([pt_name] if pt_name else [])
        with arcpy.da.SearchCursor(pt, point_fields) as cursor:
            raw_points = [tuple([row[0], projected(row[1], point_sr, pt_transform, 'point', row[0])] + list(row[2:])) for row in cursor]
        nearest = {}
        if method == BY_NEAREST and any(row[1] is not None and row[1].pointCount > 0 for row in raw_points):
            # Near analysis must use the same analysis CRS. Build scratch geometry
            # from the projected copies, retaining original source OIDs explicitly.
            suffix = uuid.uuid4().hex
            near_table = os.path.join('memory', 'near_' + suffix)
            scratch_p = os.path.join('memory', 'pt_' + suffix)
            scratch_l = os.path.join('memory', 'ln_' + suffix)
            opts = {'closest': 'ALL', 'closest_count': 2, 'method': 'PLANAR'}
            if max_distance is not None:
                opts['search_radius'] = '{} Meters'.format(max_distance)
            try:
                for path, shape in [(scratch_p, 'POINT'), (scratch_l, 'POLYLINE')]:
                    arcpy.management.CreateFeatureclass('memory', os.path.basename(path), shape, spatial_reference=sr)
                    arcpy.management.AddField(path, 'SRCID', 'TEXT', field_length=40)
                with arcpy.da.InsertCursor(scratch_p, ['SHAPE@', 'SRCID']) as cur:
                    for row in raw_points:
                        if row[1] is not None and row[1].pointCount:
                            cur.insertRow((row[1], str(row[0])))
                with arcpy.da.InsertCursor(scratch_l, ['SHAPE@', 'SRCID']) as cur:
                    for prof in profiles.values():
                        cur.insertRow((prof['geom'], prof['oid']))
                with arcpy.da.SearchCursor(scratch_p, ['OID@', 'SRCID']) as cur:
                    point_ids = {oid: src for oid, src in cur}
                with arcpy.da.SearchCursor(scratch_l, ['OID@', 'SRCID']) as cur:
                    line_ids = {oid: src for oid, src in cur}
                arcpy.AddMessage('Computing nearest profile lines in the analysis CRS.')
                arcpy.analysis.GenerateNearTable(scratch_p, scratch_l, near_table, **opts)
                with arcpy.da.SearchCursor(near_table, ['IN_FID','NEAR_FID','NEAR_DIST','NEAR_RANK']) as cur:
                    for point_oid, profile_oid, dist, rank in cur:
                        if point_oid not in point_ids or profile_oid not in line_ids:
                            raise ValueError('Nearest table references an unknown temporary feature.')
                        if dist is None or not math.isfinite(dist) or dist < 0:
                            raise ValueError('Invalid nearest-line distance for point OID {}.'.format(point_ids[point_oid]))
                        nearest.setdefault(point_ids[point_oid], []).append((int(rank), line_ids[profile_oid], float(dist)))
            finally:
                for scratch in (near_table, scratch_p, scratch_l):
                    if arcpy.Exists(scratch):
                        try:
                            arcpy.management.Delete(scratch)
                        except Exception:
                            arcpy.AddWarning('Could not remove temporary data: ' + scratch)

        points, issues = [], []
        count = {'MATCHED': 0, 'NO_NAME': 0, 'NO_PROFILE': 0, 'NO_GEOMETRY': 0, 'TIE': 0, 'NO_NEAR': 0}
        arcpy.AddMessage('Assigning observations using: {}.'.format(method))
        seen_point_oids = set()
        for row in raw_points:
                oid, geom = row[:2]
                if str(oid) in seen_point_oids:
                    raise ValueError('Duplicate point OID {}; do not use a joined layer with duplicate rows.'.format(oid))
                seen_point_oids.add(str(oid))
                name = clean(row[2]) if pt_name else ''
                if len(name) > 255:
                    raise ValueError('Point profile name exceeds 255 characters at OID {}.'.format(oid))
                if geom is None or geom.pointCount == 0:
                    count['NO_GEOMETRY'] += 1
                    issues.append((str(oid), name, 'NO_GEOMETRY'))
                    continue
                best_oid = second_oid = near1 = near2 = gap = None
                prof = None
                if method == BY_NAME:
                    k = profile_key(name)
                    status = 'NO_NAME' if not k else ('MATCHED' if k in profiles else 'NO_PROFILE')
                    if status == 'MATCHED':
                        prof = profiles[k]
                else:
                    options = sorted(nearest.get(str(oid), []), key=lambda v: (v[0], v[2], v[1]))
                    if not options:
                        status = 'NO_NEAR'
                    else:
                        _, best_oid, near1 = options[0]
                        if len(options) > 1:
                            _, second_oid, near2 = options[1]
                            gap = max(0.0, near2 - near1)
                        if len(options) > 1 and abs(near2-near1) <= TIE_TOLERANCE_M:
                            status = 'TIE'
                        else:
                            status = 'MATCHED'
                            prof = profiles_by_oid[best_oid]
                count[status] += 1
                item = {'oid': str(oid), 'geom': geom, 'name': name, 'status': status,
                        'profile_oid': None, 'group': None, 'along': None, 'offset': None,
                        'best_oid': best_oid, 'second_oid': second_oid,
                        'near1': near1, 'near2': near2, 'gap': gap}
                if status != 'MATCHED':
                    issues.append((str(oid), name, status))
                else:
                    _, along, offset, _ = prof['geom'].queryPointAndDistance(geom, False)
                    if not math.isfinite(along) or not math.isfinite(offset):
                        raise ValueError('Non-finite position or distance for point OID {}.'.format(oid))
                    if along < -1e-6 or along > prof['length'] + 1e-6:
                        raise ValueError('Invalid position along profile for point OID {}.'.format(oid))
                    along = min(prof['length'], max(0.0, along))
                    item.update(profile_oid=prof['oid'], group=prof['group'], along=along, offset=offset)
                    prof['obs'].append((along, offset))
                points.append(item)

        def addfields(path, fields):
            for name, kind, length in fields:
                if kind == 'TEXT':
                    arcpy.management.AddField(path, name, kind, field_length=length)
                else:
                    arcpy.management.AddField(path, name, kind)
        def create_fc(path, shape):
            arcpy.management.CreateFeatureclass(gdb, os.path.basename(path), shape, spatial_reference=sr)
        arcpy.AddMessage('Creating uniquely named output datasets; source layers will not be edited.')
        create_fc(paths[0], 'POLYLINE')
        addfields(paths[0], [('PROF_NAME','TEXT',255),('GROUP_VAL','TEXT',255),('SRC_OID','TEXT',40),
                            ('N_POINTS','LONG',None),('N_POS','LONG',None),('LEN_KM','DOUBLE',None),
                            ('COVERED_KM','DOUBLE',None),('COVERED_PCT','DOUBLE',None),('UNCOVERED_KM','DOUBLE',None),('UNCOVERED_PCT','DOUBLE',None),('FIRST_M','DOUBLE',None),
                            ('LAST_M','DOUBLE',None),('MAX_GAP_M','DOUBLE',None),('MAX_OFF_M','DOUBLE',None)])
        create_fc(paths[1], 'POINT')
        addfields(paths[1], [('SRC_OID','TEXT',40),('INPUT_NAME','TEXT',255),('MATCH_STATUS','TEXT',20),
                            ('PROF_OID','TEXT',40),('GROUP_VAL','TEXT',255),('ALONG_M','DOUBLE',None),('OFFSET_M','DOUBLE',None),
                            ('BEST_OID','TEXT',40),('SECOND_OID','TEXT',40),('NEAR1_M','DOUBLE',None),
                            ('NEAR2_M','DOUBLE',None),('GAP_M','DOUBLE',None)])
        create_fc(paths[2], 'POLYLINE')
        addfields(paths[2], [('PROF_NAME','TEXT',255),('GROUP_VAL','TEXT',255),('SRC_OID','TEXT',40),('COVERED_KM','DOUBLE',None)])
        arcpy.management.CreateTable(gdb, os.path.basename(paths[3]))
        addfields(paths[3], [('ROW_TYPE','TEXT',10),('GROUP_VAL','TEXT',255),('TOTAL_KM','DOUBLE',None),
                            ('COVERED_KM','DOUBLE',None),('COVERED_PCT','DOUBLE',None),('UNCOVERED_KM','DOUBLE',None),('UNCOVERED_PCT','DOUBLE',None),('N_PROFILES','LONG',None),('N_POINTS','LONG',None)])
        arcpy.management.CreateTable(gdb, os.path.basename(paths[4]))
        addfields(paths[4], [('SRC_OID','TEXT',40),('INPUT_NAME','TEXT',255),('ISSUE','TEXT',20)])

        # Group values remain case-sensitive after trimming; empty is a separate group.
        # The total is kept separate by ROW_TYPE, not by the group label.
        groups = {}
        stats_rows, span_rows = [], []
        total_length = total_span = 0.0
        for prof in profiles.values():
            obs = sorted(prof['obs']); n = len(obs); length = prof['length']
            first = obs[0][0] if n else None
            last = obs[-1][0] if n else None
            span = last - first if n >= 2 else 0.0
            gap = max((obs[i][0] - obs[i-1][0] for i in range(1, n)), default=None)
            max_offset = max((o[1] for o in obs), default=None)
            n_positions = len({o[0] for o in obs})
            stats_rows.append((prof['geom'], prof['name'], prof['group'], prof['oid'], n, n_positions,
                               length/1000, span/1000, 100*span/length, (length-span)/1000,
                               100*(length-span)/length, first, last, gap, max_offset))
            if span > 0:
                segment = prof['geom'].segmentAlongLine(first, last, False)
                span_rows.append((segment, prof['name'], prof['group'], prof['oid'], span/1000))
            total_length += length; total_span += span
            if group_name:
                s = groups.setdefault(prof['group'], {'total': 0.0, 'span': 0.0, 'nprof': 0, 'npoints': 0})
                s['total'] += length; s['span'] += span; s['nprof'] += 1; s['npoints'] += n
        with arcpy.da.InsertCursor(paths[0], ['SHAPE@','PROF_NAME','GROUP_VAL','SRC_OID','N_POINTS','N_POS','LEN_KM','COVERED_KM','COVERED_PCT','UNCOVERED_KM','UNCOVERED_PCT','FIRST_M','LAST_M','MAX_GAP_M','MAX_OFF_M']) as cur:
            for row in stats_rows: cur.insertRow(row)
        with arcpy.da.InsertCursor(paths[1], ['SHAPE@','SRC_OID','INPUT_NAME','MATCH_STATUS','PROF_OID','GROUP_VAL','ALONG_M','OFFSET_M',
                                             'BEST_OID','SECOND_OID','NEAR1_M','NEAR2_M','GAP_M']) as cur:
            for v in points:
                cur.insertRow((v['geom'],v['oid'],v['name'],v['status'],v['profile_oid'],v['group'],v['along'],v['offset'],
                               v['best_oid'],v['second_oid'],v['near1'],v['near2'],v['gap']))
        with arcpy.da.InsertCursor(paths[2], ['SHAPE@','PROF_NAME','GROUP_VAL','SRC_OID','COVERED_KM']) as cur:
            for row in span_rows: cur.insertRow(row)
        summary_fields = ['ROW_TYPE','GROUP_VAL','TOTAL_KM','COVERED_KM','COVERED_PCT','UNCOVERED_KM','UNCOVERED_PCT','N_PROFILES','N_POINTS']
        with arcpy.da.InsertCursor(paths[3], summary_fields) as cur:
            for name, s in sorted(groups.items()):
                cur.insertRow(('GROUP',name,s['total']/1000,s['span']/1000,100*s['span']/s['total'] if s['total'] else None,
                               (s['total']-s['span'])/1000,100*(s['total']-s['span'])/s['total'] if s['total'] else None,s['nprof'],s['npoints']))
            cur.insertRow(('TOTAL',None,total_length/1000,total_span/1000,
                           100*total_span/total_length if total_length else None,
                           (total_length-total_span)/1000,100*(total_length-total_span)/total_length if total_length else None,
                           len(profiles),count['MATCHED']))
        with arcpy.da.InsertCursor(paths[4], ['SRC_OID','INPUT_NAME','ISSUE']) as cur:
            for row in issues: cur.insertRow(row)
        write_excel_report(excel, groups, bool(group_name), total_length, total_span,
                           len(profiles), count['MATCHED'], stamp, sr.name,
                           str(pt), pt_name, str(line), line_name, group_name, method, max_distance,
                           [('Point source CRS', point_sr.name), ('Profile source CRS', line_sr.name),
                            ('Point transformation', pt_transform or '(none)'), ('Profile transformation', line_transform or '(none)'),
                            ('Projection caution', 'Verify the CRS and transformation for the study area. Lines are projected without automatic densification.')])

        def fmt(label, covered, length):
            return '{}: covered {:.2f} / {:.2f} km ({}); uncovered {:.2f} km ({})'.format(
                label, covered/1000, length/1000,
                '{:.1f}%'.format(100*covered/length) if length else 'N/A', (length-covered)/1000,
                '{:.1f}%'.format(100*(length-covered)/length) if length else 'N/A')
        report_lines = ['PROFILE COVERAGE - SELECTED INPUT POINTS', 'Assignment: ' + method,
                        'Covered: span between first and last matched point']
        for name, s in sorted(groups.items()):
            report_lines.append(fmt(name if name else '[Empty group]', s['span'], s['total']))
        report_lines += [fmt('Total', total_span, total_length),
                         'Estimated only; overlaps counted separately. Covered and uncovered are not confirmed traversal.',
                         'Prepared by: ' + AUTHOR]
        project_actions = []
        if map_name or layout_name:
            try:
                project = arcpy.mp.ArcGISProject('CURRENT')
                if map_name:
                    maps = [m for m in project.listMaps() if m.name == map_name]
                    if len(maps) == 1:
                        maps[0].addDataFromPath(paths[2]); project_actions.append('Span layer added to map: ' + map_name)
                    else: arcpy.AddWarning('Map update skipped: expected one exact map name, found {}.'.format(len(maps)))
                if layout_name:
                    layouts = [v for v in project.listLayouts() if v.name == layout_name]
                    if len(layouts) == 1:
                        elements = [v for v in layouts[0].listElements('TEXT_ELEMENT') if v.name == text_name]
                        if len(elements) == 1:
                            # Names are user data, not ArcGIS text formatting markup.
                            elements[0].text = '\n'.join(html.escape(s, quote=False) for s in report_lines)
                            project_actions.append('Layout text updated: {} / {}'.format(layout_name, text_name))
                            if elements[0].isOverflowing: arcpy.AddWarning('Layout text overflows: enlarge the text box or reduce the font size.')
                        else: arcpy.AddWarning('Layout update skipped: expected one exact text element, found {}.'.format(len(elements)))
                    else: arcpy.AddWarning('Layout update skipped: expected one exact layout name, found {}.'.format(len(layouts)))
                if project_actions: arcpy.AddWarning('Review map and layout; save the project manually to retain changes.')
            except Exception as e:
                arcpy.AddWarning('Optional map/layout update was not fully completed: {}'.format(e))

        lines = ['PROFILE COVERAGE - RUN NOTES', 'Prepared by: ' + AUTHOR, 'Tool version: ' + VERSION,
                 'Email: ' + EMAIL, 'LinkedIn: ' + LINKEDIN, 'GitHub: ' + GITHUB, 'Run ID: ' + stamp,
                 'Run time (local computer): ' + datetime.now().astimezone().isoformat(timespec='seconds'),
                 'Point layer: ' + str(pt), 'Point profile field: ' + (pt_name or '(not used)'),
                 'Profile line layer: ' + str(line), 'Line profile field: ' + (line_name or '(not used)'),
                 'Optional group field: ' + (group_name if group_name else '(none)'),
                 'Analysis coordinate system: ' + sr.name, 'Length: 2D planar metres; reported in km.',
                 'Point source CRS: ' + point_sr.name,
                 'Profile source CRS: ' + line_sr.name,
                 'Point transformation: ' + (pt_transform or '(none)'),
                 'Profile transformation: ' + (line_transform or '(none)'),
                 'Projection uses geometry.projectAs; line vertices are not densified to preserve projected shape. Inspect or densify long segments if required.',
                 'User must verify CRS suitability, input CRS metadata and geographic transformation area/accuracy.',
                 'Assignment method: ' + method,
                 'Name matching: Unicode NFKC, trim leading/trailing spaces, case-insensitive; no nearest-line fallback.',
                 'Nearest mode: PROFILE_<source OID> labels are generated in outputs only; source OIDs can change.',
                 'Exact equal-nearest ties (within 1e-8 m) remain unassigned; review TIE records and point audit.',
                 'Nearest/second nearest candidate OIDs and distances are logged in the point audit; review small GAP_M manually.',
                 'Group values: trimmed, case-sensitive; blanks are grouped separately.',
                 ('No maximum offset: matched points at any distance contribute; inspect OFFSET_M and MAX_OFF_M.'
                  if max_distance is None else 'Nearest-line search radius (m): {}'.format(max_distance)),
                 'COVERED is the estimated span along the profile between minimum and maximum projected point positions.',
                 'Intervening gaps remain in covered span; it is NOT proof of continuous traversal or geological coverage.',
                 'UNCOVERED is total length minus covered span; it is not confirmed untraversed distance.'
                 , 'Fewer than two points: covered span is zero; coincident positions can also give zero span.',
                 'N_POS counts distinct projected positions by exact numerical equality, not a distance tolerance.',
                 'Overlapping profiles count separately; denominator uses the full selected profile lengths.',
                 'Selections and definition queries apply; the point layer must represent the intended period.',
                 'No date filtering, polygon clipping, map symbology or project save was applied.',
                 'SRC_OID and PROF_OID refer to input layer OIDs at the time of this run; preserve source versions.',
                 'Output feature classes contain only analysis fields, not all source attributes.',
                 'Selected profile count: {}'.format(len(profiles)),
                 'Selected point count, including empty geometry: {}'.format(len(points)+count['NO_GEOMETRY']),
                 'Point status counts: ' + ', '.join('{}={}'.format(k,v) for k,v in count.items()),
                 'COVERED_PCT and UNCOVERED_PCT are 0 to 100, not Excel fractional percentages.',
                 'Check self-intersections, input quality and output visually. Do not infer area or mineral reserves.']
        for label, desc in [('Point',pdesc), ('Profile line',ldesc)]:
            lines.append(label + ' source catalog path: ' + str(getattr(desc, 'catalogPath', 'unavailable')))
        lines.extend(project_actions if project_actions else ['No optional map or layout change recorded.'])
        with open(notes_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
        if issues:
            arcpy.AddWarning('Point issues found: NO_NAME={}, NO_PROFILE={}, NO_GEOMETRY={}, TIE={}, NO_NEAR={}. Review the issues and point audit tables.'.format(
                count['NO_NAME'], count['NO_PROFILE'], count['NO_GEOMETRY'], count['TIE'], count['NO_NEAR']))
        if count['MATCHED'] == 0:
            arcpy.AddWarning('No points matched. Check inputs, selected method, search radius and active selections.')
        if max_distance is None:
            arcpy.AddWarning('No maximum offset was applied. Review distant points.')
        arcpy.AddWarning('Covered and uncovered are estimates, not confirmed traversal.')
        for index, output in enumerate(paths + [excel, notes_path], start=15):
            p[index].value = output
            arcpy.AddMessage('Output: ' + output)
        arcpy.AddMessage('Finished. Review all audit outputs before reporting the results.')


class ParameterGuide:
    """Read-only on-screen parameter guide; does not create or modify data."""
    def __init__(self):
        self.label = 'Parameter Guide | راهنمای فیلدهای ابزار'
        self.description = ('Run this help tool to display short descriptions of all input and output fields in Messages. Prepared by: ' + AUTHOR)
        self.canRunInBackground = False

    def getParameterInfo(self):
        return []

    def isLicensed(self):
        return True

    def execute(self, parameters, messages):
        arcpy.AddMessage(PARAMETER_GUIDE)
