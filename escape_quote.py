import re

input_file = "input.sql"         # Your original file
output_file = "output.sql"  # Output file ready for SQL Server

def escape_single_quotes(value):
    return value.replace("'", "''")

def split_values_block(values_block):
    """
    Splits the values in multi-row inserts into individual row tuples.
    Handles nested parentheses and quoted strings safely.
    """
    rows = []
    current = ''
    depth = 0
    in_quote = False
    prev_char = ''

    for char in values_block:
        if char == "'" and prev_char != "\\":
            in_quote = not in_quote
        if char == '(' and not in_quote:
            if depth == 0 and current:
                current = ''
            depth += 1
        if char == ')' and not in_quote:
            depth -= 1

        current += char
        if depth == 0 and current.strip():
            rows.append(current.strip().rstrip(','))
            current = ''
        prev_char = char

    return rows

with open(input_file, 'r', encoding='utf-8') as infile:
    lines = infile.readlines()

with open(output_file, 'w', encoding='utf-8') as outfile:
    for line in lines:
        if not line.strip().lower().startswith("insert into"):
            outfile.write(line)
            continue

        # Match: INSERT INTO table_name ... VALUES ...
        match = re.match(r"(INSERT INTO\s+[^\s(]+.*?VALUES\s+)(.+);?", line.strip(), re.IGNORECASE | re.DOTALL)
        if not match:
            outfile.write(line)
            continue

        insert_head = match.group(1)
        values_block = match.group(2)

        rows = split_values_block(values_block)

        for row in rows:
            # Escape single quotes inside values
            row_escaped = ''
            in_quote = False
            current = ''
            for i, char in enumerate(row):
                if char == "'" and (i == 0 or row[i - 1] != "\\"):
                    in_quote = not in_quote
                    current += char
                elif in_quote and char == "'":
                    current += "''"
                else:
                    current += char
            row_escaped = current

            outfile.write(f"{insert_head}{row_escaped};\n")

print("✅ SQL file cleaned and saved as:", output_file)