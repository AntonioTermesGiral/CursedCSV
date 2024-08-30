import { Button, Grid2 } from '@mui/material'
import writeXlsxFile, { Cell } from 'write-excel-file'

function App() {

  const generateXLSX = async (columns: string[], rows: string[][]) => {
    const col_obj = [
      ...columns.map((col) => { return { value: col }; })
    ] as Object[];

    const rows_obj = [
      ...rows.map((row) => {
        return [
          {
            type: String,
            value: row[0]
          },
          {
            type: String,
            value: row[1]
          },
          {
            type: Date,
            value: new Date(Date.parse(row[2])),
            format: 'dd/mm/yyyy',
          },
          {
            type: Number,
            value: Number.parseFloat(row[3])
          }
        ] as Object[];
      })
    ]

    const data: Object[][] = [
      col_obj,
      ...rows_obj
    ]

    await writeXlsxFile(data, { fileName: 'output.xlsx' });
  }

  const fixCSV = (rawData: ProgressEvent<FileReader>) => {
    const text = (rawData.currentTarget as FileReader).result;
    if (text) {
      const columns: string[] = [];
      const rows: string[][] = [];
      const splittedText = text.toString().split('\n');
      const rawColumn = splittedText.shift();
      const rawRows = splittedText;
      rawRows.pop();

      if (rawColumn) {
        rawColumn?.split(',').forEach((col) => {
          columns.push(col);
        })

        rawRows.forEach((row) => {
          let fixedRow = row;

          // Extra quotes handling
          fixedRow = fixedRow.split("\"\"").join("\"");
          fixedRow = fixedRow.slice(1);
          fixedRow = fixedRow.slice(0, -2);

          // Date handling
          const badDateRx = /["][^"]+[ ][^"]+[ ][^"]+["]/g; // ["].*["]
          let badDate = (fixedRow.match(badDateRx) ?? '').toString();
          const dateTimestamp = Date.parse(badDate);
          const fixedDate = new Date(dateTimestamp);
          const fixedDateStr = fixedDate.toLocaleDateString();
          fixedRow = fixedRow.replace(badDate, fixedDateStr);

          // Row mapping
          const currentRow: string[] = [];
          fixedRow.split(',').forEach((val) => {
            currentRow.push(val);
          });
          rows.push(currentRow);
        })

        console.log(columns);
        console.log(rows);
        // File generation
        generateXLSX(columns, rows).then(() => console.log("end"));

      }
    }
  }

  const handleSubmit = () => {
    const input = document.getElementById("csv-input");
    if (input) {
      const files = (input as HTMLInputElement).files;
      if (files) {
        var reader = new FileReader();
        reader.onload = fixCSV;
        reader.readAsText(files[0]);
      }
    } else {
      // TODO: show err
    }
  }

  return (
    <Grid2 container alignItems="center" justifyContent="center" mt={10} direction="column" spacing={2}>
      Sube tu archivo!
      <input id='csv-input' type='file' />
      <Button variant='outlined' onClick={handleSubmit}>Convertir a Excel</Button>
    </Grid2>
  )
}

export default App
