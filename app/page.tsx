'use client'

import { useEffect, useState } from 'react'
import * as XLSX from 'xlsx'

export default function Home() {

  const [data1, setData1] = useState<any[]>([])
  const [data2, setData2] = useState<any[]>([])

  const [headers1, setHeaders1] = useState<string[]>([])
  const [headers2, setHeaders2] = useState<string[]>([])

  const [results, setResults] = useState<any[]>([])

  const [mapping, setMapping] = useState<any>({})

  // LOAD MAPPING ĐÃ LƯU
  useEffect(() => {

    const saved =
      localStorage.getItem('excel_mapping')

    if (saved) {

      setMapping(JSON.parse(saved))

    }

  }, [])

  // LƯU MAPPING
  useEffect(() => {

    localStorage.setItem(
      'excel_mapping',
      JSON.stringify(mapping)
    )

  }, [mapping])

  async function readExcel(file: File) {

    return new Promise<any[]>((resolve) => {

      const reader = new FileReader()

      reader.onload = (e) => {

        const workbook = XLSX.read(
          e.target?.result,
          {
            type: 'binary',
          }
        )

        const sheet =
          workbook.Sheets[
            workbook.SheetNames[0]
          ]

        const json =
          XLSX.utils.sheet_to_json(sheet, {
            defval: '',
          })

        resolve(json)

      }

      reader.readAsBinaryString(file)

    })

  }

 
async function upload1(e: any) {

  const file = e.target.files[0]

  if (!file) return

  const json =
    await readExcel(file)

  setData1(json)

  if (json.length > 0) {

    const h =
      Object.keys(json[0])

    setHeaders1(h)

    // LOAD MAPPING CŨ
    const saved =
      localStorage.getItem(
        'excel_mapping'
      )

    if (saved) {

      setMapping(
        JSON.parse(saved)
      )

    }

  }

}

async function upload2(e: any) {

  const file = e.target.files[0]

  if (!file) return

  const json =
    await readExcel(file)

  setData2(json)

  if (json.length > 0) {

    const h =
      Object.keys(json[0])

    setHeaders2(h)

    // LOAD MAPPING CŨ
    const saved =
      localStorage.getItem(
        'excel_mapping'
      )

    if (saved) {

      setMapping(
        JSON.parse(saved)
      )

    }

  }

}

  function normalize(value: any) {

    return String(value ?? '')
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')

  }

  function parseVNDate(str: string) {

    const parts =
      str.split(/[\/ :]/)

    if (parts.length < 3) {
      return null
    }

    const dd = parts[0]
    const MM = parts[1]
    const yyyy = parts[2]

    const hh = parts[3] || '00'
    const mm = parts[4] || '00'

    return {
      yyyy,
      MM,
      dd,
      hh,
      mm,
    }

  }

  // FORMAT yyyyMMddHHmm
  function formatDateTime(value: any) {

    if (!value) return ''

    const str =
      String(value).trim()

    // ĐÃ ĐÚNG FORMAT
    if (/^\d{12}$/.test(str)) {
      return str
    }

    // dd/MM/yyyy HH:mm
    if (str.includes('/')) {

      const p =
        parseVNDate(str)

      if (!p) return str

      return `${p.yyyy}${p.MM}${p.dd}${p.hh}${p.mm}`

    }

    return str

  }

  // FORMAT yyyyMMdd0000
  function formatDateOnly(value: any) {

    if (!value) return ''

    const str =
      String(value).trim()

    if (/^\d{12}$/.test(str)) {
      return str
    }

    if (str.includes('/')) {

      const p =
        parseVNDate(str)

      if (!p) return str

      return `${p.yyyy}${p.MM}${p.dd}0000`

    }

    return str

  }

  // GIỚI TÍNH
  function normalizeGender(value: any) {

    const v =
      String(value ?? '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')

    if (
      v.includes('nam') ||
      v === '1' ||
      v.includes('male')
    ) {
      return '1'
    }

    if (
      v.includes('nu') ||
      v === '2' ||
      v.includes('female')
    ) {
      return '2'
    }

    return v

  }

  // KEY = MA_THE_BHYT + NGAY_RA
  function buildKey(row: any) {

    const maThe =
      normalize(
        row['MA_THE_BHYT'] ||
        row['Mã thẻ'] ||
        ''
      )

    const ngayRa =
      formatDateTime(
        row['NGAY_RA'] ||
        row['Ngày ra'] ||
        ''
      )

    return `${maThe}_${ngayRa}`

  }

  // HIỂN THỊ MAPPING
  function buildMappingInfo(
    row1: any,
    row2: any
  ) {

    const maThe1 =
      row1?.['MA_THE_BHYT'] || ''

    const ngayRa1 =
      formatDateTime(
        row1?.['NGAY_RA'] || ''
      )

    const maThe2 =
      row2?.['Mã thẻ'] || ''

    const ngayRa2 =
      formatDateTime(
        row2?.['Ngày ra'] || ''
      )

    return `${maThe1} | ${ngayRa1} ↔ ${maThe2} | ${ngayRa2}`

  }

  function processValue(
    columnName: string,
    value: any
  ) {

    const col =
      normalize(columnName)

    // GIỚI TÍNH
    if (
      col.includes('gioi') ||
      col.includes('tinh') ||
      col.includes('gender')
    ) {

      return normalizeGender(value)

    }

    // NGÀY SINH
    if (
      col.includes('ngay_sinh') ||
      col.includes('ngay sinh')
    ) {

      return formatDateOnly(value)

    }

    // NGÀY VÀO / RA
    if (
      col.includes('ngay_vao') ||
      col.includes('ngay vao') ||
      col.includes('ngay_ra') ||
      col.includes('ngay ra')
    ) {

      return formatDateTime(value)

    }

    // TIỀN / SỐ
    const raw =
      String(value)
        .replace(/,/g, '')
        .trim()

    const num =
      Number(raw)

    if (!isNaN(num)) {

      return num.toFixed(2)

    }

    return normalize(value)

  }


function compare() {

  const diff: any[] = []

  const map2 = new Map()

  // MAP FILE PHẢI
  data2.forEach((r2) => {

    const key =
      buildKey(r2)

    map2.set(key, r2)

  })

  // DUYỆT FILE TRÁI
  data1.forEach((r1, index) => {

    const key =
      buildKey(r1)

    const r2 =
      map2.get(key)

    const stt =
      r1['STT'] ||
      index + 1

    const mappingInfo =
      buildMappingInfo(
        r1,
        r2 || {}
      )

    // KHÔNG TỒN TẠI
    if (!r2) {

      diff.push({
        row: stt,
        mapping: mappingInfo,
        column: 'KEY',
        file1: key,
        file2: 'Không tồn tại',
      })

      return

    }

    // SO SÁNH CỘT
    Object.entries(mapping).forEach(
      ([col1, col2]) => {

        const c1 =
          String(col1)

        const c2 =
          String(col2)

        if (!c2) return

        // KHÔNG SO SÁNH STT
        if (
          normalize(c1) === 'stt'
        ) {
          return
        }

        const v1 =
          processValue(
            c1,
            r1[c1]
          )

        const v2 =
          processValue(
            c2,
            r2[c2]
          )

        if (v1 !== v2) {

          diff.push({
            row: stt,
            mapping: mappingInfo,
            column: `${c1} ↔ ${c2}`,
            file1: v1,
            file2: v2,
          })

        }

      }
    )

  })

  // SORT DÒNG
  diff.sort((a, b) => {

    return (
      Number(a.row) -
      Number(b.row)
    )

  })

  setResults(diff)

}

  function exportExcel() {

    const ws =
      XLSX.utils.json_to_sheet(results)

    const wb =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      wb,
      ws,
      'KetQua'
    )

    XLSX.writeFile(
      wb,
      'doi_chieu.xlsx'
    )

  }

  return (

    <main className="min-h-screen bg-gray-100 p-8">

      <div className="max-w-7xl mx-auto bg-white rounded-2xl shadow-xl p-6">

        <h1 className="text-4xl font-bold mb-8 text-center">
          ĐỐI CHIẾU EXCEL BHYT
        </h1>

        <div className="grid grid-cols-2 gap-6 mb-8">

          {/* FILE TRÁI */}
          <div className="bg-blue-50 border-2 border-blue-300 rounded-2xl p-6">

            <h2 className="text-xl font-bold mb-4 text-blue-700">
              FILE BÊN TRÁI
            </h2>

            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={upload1}
            />

            <div className="mt-4 text-sm">
              Số dòng: {data1.length}
            </div>

          </div>

          {/* FILE PHẢI */}
          <div className="bg-green-50 border-2 border-green-300 rounded-2xl p-6">

            <h2 className="text-xl font-bold mb-4 text-green-700">
              FILE BÊN PHẢI
            </h2>

            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={upload2}
            />

            <div className="mt-4 text-sm">
              Số dòng: {data2.length}
            </div>

          </div>

        </div>

        {/* MAPPING */}
        <div className="bg-yellow-50 border rounded-2xl p-6 mb-8">

          <h2 className="text-2xl font-bold mb-4">
            Mapping cột
          </h2>

          <div className="grid grid-cols-2 gap-4">

            {headers1.map((h1) => (

              <div
                key={h1}
                className="flex gap-2 items-center"
              >

                <div className="w-1/2 text-sm font-medium">
                  {h1}
                </div>

                <select
                  className="border rounded-lg p-2 w-1/2"
                  value={mapping[h1] || ''}
                  onChange={(e) => {

                    setMapping({
                      ...mapping,
                      [h1]: e.target.value,
                    })

                  }}
                >

                  <option value="">
                    -- Chọn cột --
                  </option>

                  {headers2.map((h2) => (

                    <option
                      key={h2}
                      value={h2}
                    >
                      {h2}
                    </option>

                  ))}

                </select>

              </div>

            ))}

          </div>

        </div>

        {/* BUTTON */}
        <div className="flex gap-4 mb-8">

          <button
            onClick={compare}
            className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-3 rounded-xl font-bold"
          >
            Đối chiếu
          </button>

          <button
            onClick={exportExcel}
            className="bg-green-600 hover:bg-green-700 text-white px-8 py-3 rounded-xl font-bold"
          >
            Export Excel
          </button>

        </div>

        {/* KẾT QUẢ */}
        <div className="mb-4 text-lg font-bold">
          Tổng lỗi: {results.length}
        </div>

        <div className="overflow-auto max-h-[700px] border rounded-xl">

          <table className="w-full border-collapse text-sm">

            <thead className="sticky top-0 bg-gray-200">

              <tr>

                <th className="border p-2">
                  Dòng
                </th>

                <th className="border p-2">
                  Mapping
                </th>

                <th className="border p-2">
                  Cột
                </th>

                <th className="border p-2">
                  File trái
                </th>

                <th className="border p-2">
                  File phải
                </th>

              </tr>

            </thead>

            <tbody>

              {results.map((r, i) => (

                <tr
                  key={i}
                  className="bg-red-50"
                >

                  <td className="border p-2">
                    {r.row}
                  </td>

                  <td className="border p-2">
                    {r.mapping}
                  </td>

                  <td className="border p-2">
                    {r.column}
                  </td>

                  <td className="border p-2">
                    {r.file1}
                  </td>

                  <td className="border p-2">
                    {r.file2}
                  </td>

                </tr>

              ))}

            </tbody>

          </table>

        </div>

      </div>

    </main>

  )

}