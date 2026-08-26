import re
import os

class Integrator:
    class Tee:
        def __init__(self, *files):
            self.files = files
        def write(self, obj):
            for f in self.files:
                f.write(obj)
                f.flush()
        def flush(self):
            for f in self.files:
                f.flush()

    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.output_path = output_path

    def parse_filename(fname):
        name = os.path.splitext(fname)[0]
        pattern = re.compile(
            r"""
            ^(?P<country>.+?)       # Non-greedy match for country name
            (?P<company>\d+)?       # Optional digit(s) for company number
            -
            (?P<start>\d{4})        # 4-digit start year
            (?:-(?P<end>\d{4}))?    # Optional hyphen and 4-digit end year
            (?P<suffix>[A-Za-z]+)$  # Suffix variable(s), e.g., "A" or "AB"
            """,
            re.VERBOSE
        )
        
        m = pattern.fullmatch(name)
        if not m:
            return None
            
        info = m.groupdict()
        
        return {
            'country': info['country'],
            'company': int(info['company']) if info['company'] else None,
            'start': int(info['start']),
            'end': int(info['end']) if info['end'] else None,
            'suffix': list(info['suffix']) if len(info['suffix']) > 1 else info['suffix']
        }

    def actual_rows(ws):
        last = 0
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
            if any(cell is not None for cell in row):
                last = i
        return last

    def actual_cols(ws):
        max_cols = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            for i in range(len(row), 0, -1):
                if row[i-1] is not None:
                    max_cols = max(max_cols, i)
                    break
        return max_cols
