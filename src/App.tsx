import { useEffect, useRef, useState } from 'react'
import { Workbook } from 'exceljs';
import './App.css'

function App() {
  const [file, setFile] = useState<File | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const handleUpload = (f: File) => {
    const fileReader = new FileReader();
    fileReader.onload = (e) => {
      const data = e.target?.result;
      if (data) {
        const arrayBuffer = data as ArrayBuffer;
        const blob = new Blob([arrayBuffer], { type: f.type });
        const imageUrl = URL.createObjectURL(blob);
        const img = new Image();
        img.src = imageUrl;
        img.onload = () => {
          const canvas = canvasRef.current;
            const ctx = canvas!.getContext('2d');
            ctx!.imageSmoothingEnabled = false;
            canvas!.width = 300;
            canvas!.height = 300 * (img.height / img.width);
            ctx!.drawImage(img, 0, 0, canvas!.width, canvas!.height);
        }
      }
    }
    fileReader.readAsArrayBuffer(f);
  }

  const convertToExcel = () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Result");
    const canvas = canvasRef.current;
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (ctx) {
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imageData.data;
        const rowCount = Math.ceil(canvas.height);
        const colCount = Math.ceil(canvas.width);
        for (let c = 0; c < colCount; c++) {
          const col = worksheet.getColumn(c + 1);
          col.width = 2;
        }
        for (let r = 0; r < rowCount; r++) {
          const row = worksheet.getRow(r + 1);
          for (let c = 0; c < colCount; c++) {
            const index = (r * colCount + c) * 4;
            if (index < data.length) {
              const color = `#${data[index].toString(16).padStart(2, '0')}${data[index + 1].toString(16).padStart(2, '0')}${data[index + 2].toString(16).padStart(2, '0')}`;
              row.getCell(c + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color.replace('#', '') } };
            }
          }
        }
      }
    }
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer]);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = file?.name.replace(/\.[^/.]+$/, "") + '.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    });
  }

  useEffect(() => {
    if (file) {
      handleUpload(file);
    }
  }, [file]);
  return (
    <>
      <div className="container py-5">
        <div className="h3 fw-bold">XLSX - Image Converter</div>
        <div className='fst-italic'>By <span className='fw-semibold'>@qxbao</span></div>
        <div className="my-4">
          <div className='mb-2'>Select input file</div>
          <div className="d-flex">
            <input
              draggable
              type="file"
              className="flex-shrink-1 form-control w-auto"
              accept="image/*"
              onChange={(e) => {
                if (e.target.files && e.target.files.length > 0) {
                  setFile(e.target.files[0])
                } else {
                  setFile(null)
                }
              }}
            />
          </div>
        </div>
        <div>
          <canvas ref={canvasRef} className=''>
          </canvas>
        </div>
        <div>
          <button className="btn btn-primary"
            disabled={!file}
            onClick={convertToExcel}
          >
            Convert file
          </button>
        </div>
      </div>
    </>
  )
}

export default App
