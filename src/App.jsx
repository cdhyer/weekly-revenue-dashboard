import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

function Card({ children, className = "" }) {
  return (
    <div className={`bg-white border border-neutral-200 rounded-2xl ${className}`}>
      {children}
    </div>
  );
}

function CardContent({ children, className = "" }) {
  return <div className={className}>{children}</div>;
}

export default function WeeklyRevenueDashboard() {
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [selectedDateKey, setSelectedDateKey] = useState("");
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");

  const money = (value) => {
    const n = Number(value || 0);
    const formatted = Math.abs(n).toLocaleString(undefined, {
      style: "currency",
      currency: "USD",
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
    return n < 0 ? `-${formatted}` : formatted;
  };

  const normalize = (value) => String(value ?? "").trim().toLowerCase();

  const excelDateToDate = (value) => {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (!parsed) return null;
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }

    const text = String(value ?? "").trim();
    if (!text) return null;

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  };

  const dateKey = (date) => {
    if (!date) return "";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(
      2,
      "0"
    )}-${String(date.getDate()).padStart(2, "0")}`;
  };

  const prettyDate = (date) => {
    if (!date) return "";
    return date.toLocaleDateString(undefined, {
      month: "numeric",
      day: "numeric",
      year: "2-digit",
    });
  };

  const toNumber = (value) => {
    if (typeof value === "number") return value;
    if (value == null || value === "") return 0;

    const cleaned = String(value).replace(/[$,()]/g, "").trim();
    if (!cleaned) return 0;

    const n = Number(cleaned);
    if (Number.isNaN(n)) return 0;

    return String(value).includes("(") ? -n : n;
  };

  const handleUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setError("");
    setFileName(file.name);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
      });

      const sheet = workbook.Sheets["2026"];

      if (!sheet) {
        setError("Could not find a tab named '2026'.");
        return;
      }

      const data = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: true,
      });

      const headerRowIndex = data.findIndex(
        (row) =>
          row.some((cell) => normalize(cell) === "date") &&
          row.some((cell) => normalize(cell) === "days out")
      );

      if (headerRowIndex === -1) {
        setError("Could not find a header row with 'date' and 'days out'.");
        return;
      }

      const headerRow = data[headerRowIndex].map((h) => String(h ?? "").trim());
      const dateIndex = headerRow.findIndex((h) => normalize(h) === "date");
      const daysOutIndex = headerRow.findIndex(
        (h) => normalize(h) === "days out"
      );

      const parsedRows = data
        .slice(headerRowIndex + 1)
        .map((row) => {
          const date = excelDateToDate(row[dateIndex]);
          const daysOut = Number(row[daysOutIndex]);

          if (!date || Number.isNaN(daysOut)) return null;

          return {
            raw: row,
            date,
            dateKey: dateKey(date),
            year: date.getFullYear(),
            daysOut,
          };
        })
        .filter(Boolean);

      setHeaders(headerRow);
      setRows(parsedRows);

      const first2026 = parsedRows.find((row) => row.year === 2026);
      setSelectedDateKey(first2026?.dateKey || "");
    } catch (err) {
      setError(`Upload failed: ${err.message}`);
    }
  };

  const dashboard = useMemo(() => {
    if (!rows.length || !headers.length || !selectedDateKey) return null;

    const selected = rows.find(
      (row) => row.dateKey === selectedDateKey && row.year === 2026
    );

    if (!selected) return null;

    const categoryIndexes = [];

    // Columns C through AP
    for (let i = 2; i <= 41; i += 1) {
      if (headers[i]) categoryIndexes.push(i);
    }

    // Column AQ
    const totalIndex = 42;

    const years = [...new Set(rows.map((row) => row.year))]
      .filter((year) => year < 2026)
      .sort((a, b) => b - a);

    const findRow = (year, daysOut) =>
      rows.find(
        (row) => row.year === year && Number(row.daysOut) === Number(daysOut)
      );

    const buildCategory = (index) => {
      const currentTotal = toNumber(selected.raw[index]);

      return {
        name: headers[index],
        currentTotal,
        comparisons: years.map((year) => {
          const sameDaysRow = findRow(year, selected.daysOut);
          const sameDaysTotal = sameDaysRow
            ? toNumber(sameDaysRow.raw[index])
            : null;
          const difference = sameDaysRow ? currentTotal - sameDaysTotal : null;

          const forwardWeeks = [1, 2, 3, 4].map((week) => {
            const startDays = selected.daysOut - (week - 1) * 7;
            const endDays = selected.daysOut - week * 7;

            const startRow = week === 1 ? sameDaysRow : findRow(year, startDays);
            const endRow = findRow(year, endDays);

            if (!startRow || !endRow) return null;

            return toNumber(endRow.raw[index]) - toNumber(startRow.raw[index]);
          });

          return {
            year,
            difference,
            forwardWeeks,
          };
        }),
      };
    };

    return {
      selected,
      categories: categoryIndexes.map(buildCategory),
      totals: headers[totalIndex] ? buildCategory(totalIndex) : null,
    };
  }, [rows, headers, selectedDateKey]);

  const dates2026 = useMemo(
    () =>
      rows
        .filter((row) => row.year === 2026)
        .sort((a, b) => a.date - b.date),
    [rows]
  );

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#f5f7fa",
        padding: "32px",
        color: "#111827",
        fontFamily:
          "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif",
      }}
    >
      <div style={{ maxWidth: "1500px", margin: "0 auto" }}>
        <div style={{ marginBottom: "28px" }}>
          <h1 style={{ fontSize: "36px", fontWeight: "900", margin: 0 }}>
            Weekly Revenue Dashboard
          </h1>
          <p style={{ marginTop: "8px", color: "#5f6673", fontSize: "16px" }}>
            Upload the weekly workbook and select a 2026 date to compare category
            totals by matching days out.
          </p>
        </div>

        <Card className="shadow-sm">
          <CardContent className="space-y-4 p-5">
            <div
              style={{
                display: "flex",
                gap: "24px",
                alignItems: "flex-end",
                justifyContent: "space-between",
                flexWrap: "wrap",
              }}
            >
              <div>
                <label style={{ display: "block", fontWeight: "800" }}>
                  Upload workbook
                </label>
                <input
                  style={{
                    marginTop: "8px",
                    display: "block",
                    width: "100%",
                    border: "1px solid #cfd6df",
                    borderRadius: "10px",
                    background: "white",
                    padding: "10px",
                    fontSize: "15px",
                  }}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleUpload}
                />
                {fileName && (
                  <p style={{ marginTop: "8px", fontSize: "13px", color: "#667085" }}>
                    Loaded: {fileName}
                  </p>
                )}
              </div>

              <div>
                <label style={{ display: "block", fontWeight: "800" }}>
                  Select 2026 date
                </label>
                <select
                  style={{
                    marginTop: "8px",
                    minWidth: "240px",
                    border: "1px solid #cfd6df",
                    borderRadius: "10px",
                    background: "white",
                    padding: "10px",
                    fontSize: "15px",
                  }}
                  value={selectedDateKey}
                  onChange={(event) => setSelectedDateKey(event.target.value)}
                  disabled={!dates2026.length}
                >
                  {dates2026.map((row) => (
                    <option key={row.dateKey} value={row.dateKey}>
                      {prettyDate(row.date)}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {error && (
              <div
                style={{
                  borderRadius: "10px",
                  background: "#fff1f0",
                  color: "#b3261e",
                  padding: "12px",
                  fontSize: "15px",
                  fontWeight: "700",
                }}
              >
                {error}
              </div>
            )}
          </CardContent>
        </Card>

        {dashboard && (
          <>
            <div
              style={{
                marginTop: "24px",
                marginBottom: "24px",
                borderRadius: "18px",
                background: "white",
                padding: "24px",
                border: "1px solid #d7dce2",
                boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
              }}
            >
              <div
                style={{
                  fontSize: "13px",
                  textTransform: "uppercase",
                  letterSpacing: "0.08em",
                  color: "#667085",
                  fontWeight: "800",
                }}
              >
                Selected week
              </div>
              <div style={{ marginTop: "4px", fontSize: "30px", fontWeight: "900" }}>
                {prettyDate(dashboard.selected.date)}
              </div>
              <div style={{ marginTop: "4px", fontSize: "20px" }}>
                Days out:{" "}
                <span style={{ fontWeight: "900" }}>
                  {dashboard.selected.daysOut}
                </span>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(760px, 1fr))",
                gap: "28px",
                alignItems: "start",
              }}
            >
              {dashboard.categories.map((category) => (
                <CategoryCard
                  key={category.name}
                  category={category}
                  money={money}
                />
              ))}
            </div>

            {dashboard.totals && (
              <div style={{ marginTop: "36px" }}>
                <h2 style={{ fontSize: "32px", fontWeight: "900" }}>Totals</h2>
                <CategoryCard
                  category={dashboard.totals}
                  money={money}
                  featured
                />
              </div>
            )}
          </>
        )}

        {!dashboard && !error && (
          <div
            style={{
              marginTop: "24px",
              borderRadius: "18px",
              border: "2px dashed #cfd6df",
              background: "white",
              padding: "48px",
              textAlign: "center",
              color: "#667085",
              fontWeight: "700",
            }}
          >
            Upload an Excel workbook to begin.
          </div>
        )}
      </div>
    </div>
  );
}

function CategoryCard({ category, money, featured = false }) {
  const amountStyle = (amount, isComparison = false) => {
    if (amount == null) return { color: "#999" };

    if (amount > 0) {
      return {
        color: "#137333",
        fontWeight: isComparison ? "900" : "700",
      };
    }

    if (amount < 0) {
      return {
        color: "#b3261e",
        fontWeight: isComparison ? "900" : "700",
      };
    }

    return {
      color: "#222",
      fontWeight: isComparison ? "800" : "600",
    };
  };

  return (
    <div
      style={{
        background: "white",
        border: featured ? "2px solid #111827" : "1px solid #d7dce2",
        borderRadius: "18px",
        overflow: "hidden",
        boxShadow: "0 2px 10px rgba(0,0,0,0.06)",
      }}
    >
      <div
        style={{
          textAlign: "center",
          padding: "24px 16px",
          borderBottom: "1px solid #d7dce2",
          background: "#f8fafc",
        }}
      >
        <div
          style={{
            fontSize: "28px",
            fontWeight: "900",
            letterSpacing: "0.04em",
            marginBottom: "8px",
          }}
        >
          {category.name}
        </div>

        <div
          style={{
            fontSize: "22px",
            fontWeight: "900",
          }}
        >
          2026: {money(category.currentTotal)}
        </div>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            minWidth: "760px",
            fontSize: "18px",
          }}
        >
          <thead>
            <tr style={{ background: "#f1f3f5" }}>
              {[
                "Year",
                "Vs. 2026",
                "1 Week Forward",
                "2 Weeks Forward",
                "3 Weeks Forward",
                "4 Weeks Forward",
              ].map((heading) => (
                <th
                  key={heading}
                  style={{
                    padding: "16px 14px",
                    border: "1px solid #d7dce2",
                    textAlign: "center",
                    fontWeight: "900",
                    whiteSpace: "nowrap",
                  }}
                >
                  {heading}
                </th>
              ))}
            </tr>
          </thead>

          <tbody>
            {category.comparisons.map((row) => (
              <tr key={row.year}>
                <td
                  style={{
                    padding: "18px 14px",
                    border: "1px solid #d7dce2",
                    textAlign: "center",
                    fontWeight: "900",
                    background: "#fbfbfc",
                  }}
                >
                  {row.year}
                </td>

                <td
                  style={{
                    padding: "18px 14px",
                    border: "1px solid #d7dce2",
                    textAlign: "center",
                    ...amountStyle(row.difference, true),
                  }}
                >
                  {row.difference == null ? "-" : money(row.difference)}
                </td>

                {row.forwardWeeks.map((amount, index) => (
                  <td
                    key={index}
                    style={{
                      padding: "18px 14px",
                      border: "1px solid #d7dce2",
                      textAlign: "center",
                      ...amountStyle(amount),
                    }}
                  >
                    {amount == null ? "-" : money(amount)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}