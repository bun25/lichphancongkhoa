const { google } = require("googleapis");
const { readFileSync } = require("fs");
const express = require("express");
const app = express();

app.use(express.json());

const credentials = JSON.parse(
    readFileSync("c:\\quan ly thoi gian\\evident-lock-359511-1db86211f7b6.json", "utf8")
);

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: SCOPES,
});

const sheets = google.sheets({ version: "v4", auth });

const SPREADSHEET_ID = "1ViN7zCpcYRZki_bMVxbeV7LJjC4wUIh9YFVj_dD7i6E";

// Hàm đọc dữ liệu từ Google Sheets
async function readSheet(sheetName, range) {
    console.log(`Đang đọc dữ liệu từ sheet "${sheetName}", phạm vi "${range}"`);
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!${range}`,
    });
    const values = response.data.values || [];
    console.log(`Dữ liệu đọc được từ sheet "${sheetName}":`, values);

    // Chỉ kiểm tra định dạng thời gian ở cột B, ngoại trừ các sheet "nhân viên", "thủ thuật", "mẫu", "kết quả"
    if (!["nhân viên", "thủ thuật", "mẫu", "kết quả"].includes(sheetName)) {
        values.forEach((row, rowIndex) => {
            const cell = row[1]; // Cột B là cột thứ 2 (chỉ số 1)
            try {
                if (typeof cell === "string" && cell.trim() !== "") {
                    console.log(`Đang kiểm tra thời gian tại sheet "${sheetName}", ô B${rowIndex + 1}: "${cell}"`);
                    convertTo24HourFormat(cell); // Kiểm tra định dạng thời gian
                }
            } catch (error) {
                console.error(
                    `Dữ liệu không hợp lệ tại sheet "${sheetName}", ô B${rowIndex + 1}: "${cell}". Lỗi: ${error.message}`
                );
            }
        });
    }

    return values;
}

// Hàm ghi dữ liệu lên Google Sheets
async function writeSheet(sheetName, range, values) {
    await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!${range}`,
        valueInputOption: "RAW",
        requestBody: {
            values,
        },
    });
}

// Hàm lấy dữ liệu từ cột A của sheet nhân viên
async function getSheetNamesFromColumn(sheetName) {
    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A:A`, // Lấy toàn bộ cột A
    });
    return response.data.values.flat(); // Trả về danh sách tên sheet
}

// Hàm hỗ trợ: Nhân bản một sheet mẫu
async function duplicateSheet(templateSheetName, newSheetName) {
    console.log(`Đang nhân bản sheet mẫu: ${templateSheetName} thành sheet mới: ${newSheetName}`);
    const sheetsData = (await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID })).data.sheets;
    const templateSheet = sheetsData.find(sheet => sheet.properties.title === templateSheetName);

    if (!templateSheet) {
        throw new Error(`Sheet mẫu "${templateSheetName}" không tồn tại.`);
    }

    // Tạo sheet mới bằng cách nhân bản sheet mẫu
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
            requests: [
                {
                    duplicateSheet: {
                        sourceSheetId: templateSheet.properties.sheetId,
                        newSheetName: newSheetName,
                    },
                },
            ],
        },
    });
    console.log(`Đã nhân bản sheet mẫu: ${templateSheetName} thành sheet mới: ${newSheetName}`);
}

// Hàm tạo các sheet mới từ cột A của sheet nhân viên bằng cách nhân bản sheet mẫu
async function createSheetsFromTemplate(employeeSheetName, templateSheetName) {
    const sheetNames = await getSheetNamesFromColumn(employeeSheetName);

    // Lấy danh sách các sheet đã tồn tại trong file Google Sheets
    const existingSheets = (await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID }))
        .data.sheets.map(s => s.properties.title);

    for (const sheetName of sheetNames) {
        if (existingSheets.includes(sheetName)) {
            console.log(`⚠️ Sheet "${sheetName}" đã tồn tại. Bỏ qua.`);
            continue; // Bỏ qua nếu đã có
        }

        try {
            await duplicateSheet(templateSheetName, sheetName);
        } catch (error) {
            console.error(`❌ Lỗi khi nhân bản sheet "${sheetName}": ${error.message}`);
        }
    }
}

// Hàm xử lý chia giờ thực hiện thủ thuật
async function allocateProcedures() {
    console.log("Bắt đầu xử lý chia giờ thực hiện thủ thuật...");
    // 🔥 XÓA TOÀN BỘ DỮ LIỆU CŨ TRONG SHEET "kết quả"
    await writeSheet("kết quả", "A1:Z1000", [[""]]);
    const employees = await readSheet("nhân viên", "A:B");
    const procedures = await readSheet("thủ thuật", "A:D");

    // Chuẩn hóa dữ liệu thủ thuật
    const procedureMap = {};
    procedures.forEach(([name, type, total, direct]) => {
        procedureMap[name] = {
            type,
            totalTime: parseInt(total),
            directTime: parseInt(direct)
        };
    });

    const WORK_HOURS = {
        morningStart: parseTime("07:00"),
        morningEnd: parseTime("11:30"),
        afternoonStart: parseTime("13:00"),
        afternoonEnd: parseTime("16:30"),
        eveningStart: parseTime("16:30")
    };

    // Lịch từng nhân viên + metadata
    const employeeSchedules = {};
    employees.forEach(([name, type]) => {
        employeeSchedules[name] = {
            type,
            schedule: [],
            overtime: 0,
            patients: new Set(),
            overtimeSlots: []
        };
    });

    const sheetsData = (await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID })).data.sheets;
    const patientSheets = sheetsData.map(s => s.properties.title).filter(title => !["nhân viên", "thủ thuật", "kết quả", "mẫu"].includes(title));

    const results = [];
    const assignedProceduresPerPatient = {};

    for (const sheetName of patientSheets) {
        const patients = await readSheet(sheetName, "A:C");

        for (const [patientName, startTimeStr, procedureListStr] of patients) {
            if (!patientName || !startTimeStr || !procedureListStr) continue;

            let currentTime = parseTime(convertTo24HourFormat(startTimeStr)) + 1;
            const procedureNames = procedureListStr.split(",").map(p => p.trim());

            for (const procedureName of procedureNames) {
                if (procedureName === "Thủy châm") continue; // BỎ QUA thủ thuật Thủy châm

                const proc = procedureMap[procedureName];
                if (!proc) continue;
                const { type, totalTime, directTime } = proc;

                // Danh sách nhân viên phù hợp chuyên môn
                const qualified = employees.filter(e => e[1] === type).map(e => e[0]);

                // Ưu tiên: (1) trùng tên sheet, (2) người đã làm cho bệnh nhân, (3) ít ngoài giờ
                const preferred = qualified.sort((a, b) => {
                    const ea = employeeSchedules[a];
                    const eb = employeeSchedules[b];
                    const aScore = (a === sheetName ? -100 : 0) + (ea.patients.has(patientName) ? -10 : 0) + ea.overtime;
                    const bScore = (b === sheetName ? -100 : 0) + (eb.patients.has(patientName) ? -10 : 0) + eb.overtime;
                    return aScore - bScore;
                });

                let assigned = false;

                for (const empName of preferred) {
                    const emp = employeeSchedules[empName];
                    const lastEnd = emp.schedule.length ? emp.schedule[emp.schedule.length - 1].end : WORK_HOURS.morningStart;
                    let availableTime = Math.max(lastEnd + 1, currentTime); // +1 phút nghỉ

                    // Nếu đang trong giờ nghỉ trưa, đẩy sang chiều
                    if (availableTime >= WORK_HOURS.morningEnd && availableTime < WORK_HOURS.afternoonStart) {
                        availableTime = WORK_HOURS.afternoonStart;
                    }

                    const endTime = availableTime + directTime;
                    const outOfHours = (endTime > WORK_HOURS.morningEnd && availableTime < WORK_HOURS.afternoonStart) || endTime > WORK_HOURS.afternoonEnd;

                    // Nếu đang trong giờ làm việc hoặc chấp nhận ngoài giờ hợp lý
                    if (!outOfHours || emp.overtime <= 60) {
                        emp.schedule.push({ patientName, procedureName, start: availableTime, end: endTime });
                        emp.overtime += outOfHours ? directTime : 0;
                        if (outOfHours) {
                            emp.overtimeSlots.push(`${formatTime(availableTime)} - ${formatTime(endTime)}`);
                        }
                        emp.patients.add(patientName);
                        results.push({ patientName, empName, procedureName, start: availableTime, end: availableTime + totalTime, direct: directTime, total: totalTime, isOvertime: outOfHours });

                        if (!assignedProceduresPerPatient[patientName]) assignedProceduresPerPatient[patientName] = [];
                        assignedProceduresPerPatient[patientName].push(procedureName);

                        currentTime = availableTime + totalTime + 1; // +1 phút nghỉ
                        assigned = true;
                        break;
                    }
                }
            }
        }
    }

    // Ghi kết quả vào sheet "kết quả"
    const output = [["Tên bệnh nhân", "Tên nhân viên", "Thủ thuật", "Thời gian bắt đầu", "Thời gian kết thúc", "Loại giờ", "Thời gian trực tiếp", "Tổng thời gian"]];
    results.forEach(r => {
        output.push([
            r.patientName,
            r.empName,
            r.procedureName,
            formatTime(r.start),
            formatTime(r.end),
            r.isOvertime ? "Ngoài giờ" : "Hành chính",
            `${r.direct} phút`,
            `${r.total} phút`
        ]);
    });
    await writeSheet("kết quả", "A1", output);

    // Thống kê slot rảnh ≥10 phút
    const slotOutput = [["Tên nhân viên", "Slot thời gian rảnh (≥10 phút)"]];
    for (const [empName, empData] of Object.entries(employeeSchedules)) {
        const { schedule } = empData;
        const freeSlots = [];
        let lastEnd = WORK_HOURS.morningStart;

        for (const s of schedule) {
            if (s.start - lastEnd >= 10) {
                freeSlots.push(`${formatTime(lastEnd)} - ${formatTime(s.start)}`);
            }
            lastEnd = s.end + 1;
        }

        if (lastEnd < WORK_HOURS.afternoonEnd && WORK_HOURS.afternoonEnd - lastEnd >= 10) {
            freeSlots.push(`${formatTime(lastEnd)} - ${formatTime(WORK_HOURS.afternoonEnd)}`);
        }

        slotOutput.push([empName, freeSlots.join(", ")]);
    }
    await writeSheet("kết quả", "J1", slotOutput);

    // Thống kê tổng thời gian và chi tiết làm ngoài giờ
    const overtimeSummary = [["Tên nhân viên", "Tổng thời gian làm ngoài giờ (phút)", "Khoảng thời gian ngoài giờ"]];
    for (const [empName, empData] of Object.entries(employeeSchedules)) {
        overtimeSummary.push([empName, `${empData.overtime} phút`, empData.overtimeSlots.join(", ")]);
    }
    await writeSheet("kết quả", "M1", overtimeSummary);

    // Kiểm tra các bệnh nhân còn thiếu thủ thuật chưa được gán
    const missingAssignments = [["Bệnh nhân", "Sheet", "Thủ thuật bị thiếu"]];
    for (const sheetName of patientSheets) {
        const patients = await readSheet(sheetName, "A:C");
        for (const [patientName, , procedureListStr] of patients) {
            if (!patientName || !procedureListStr) continue;
            const expected = procedureListStr.split(",").map(p => p.trim());
            const assigned = assignedProceduresPerPatient[patientName] || [];
            const missing = expected.filter(p => !assigned.includes(p));
            if (missing.length > 0) {
                missingAssignments.push([patientName, sheetName, missing.join(", ")]);
                console.warn(`⚠️ Bệnh nhân ${patientName} (sheet ${sheetName}) thiếu thủ thuật: ${missing.join(", ")}`);
            }
        }
    }
    await writeSheet("kết quả", "P1", missingAssignments);

// Gửi dữ liệu phân công về từng sheet bệnh nhân (ghi từ cột D1 trở đi)
const patientAssignmentsBySheet = {};

// Gom dữ liệu phân công theo từng sheet
for (const r of results) {
    if (!patientAssignmentsBySheet[r.patientName]) {
        patientAssignmentsBySheet[r.patientName] = [];
    }
    const duration = r.end - r.start;
    patientAssignmentsBySheet[r.patientName].push([
        r.patientName,
        formatTime(r.start),
        r.procedureName,
        r.empName,
        formatTime(r.start),
        formatTime(r.end),
        r.isOvertime ? "Ngoài giờ" : "Hành chính",
        `${duration} phút`
    ]);
}

// Ghi vào từng sheet chứa bệnh nhân
for (const sheetName of patientSheets) {
    const sheetPatients = await readSheet(sheetName, "A:A");
    const patientNamesInSheet = sheetPatients.map(row => row[0]);

    const matchingAssignments = [];
    for (const patientName of patientNamesInSheet) {
        if (patientAssignmentsBySheet[patientName]) {
            matchingAssignments.push(...patientAssignmentsBySheet[patientName]);
        }
    }

    if (matchingAssignments.length > 0) {
        // 🔥 Xoá dữ liệu cũ từ cột D trở về sau trước khi ghi mới
        await writeSheet(sheetName, "D1:Z1000", [[""]]);

        await writeSheet(sheetName, "D1", [
            ["Tên bệnh nhân", "Thời gian bắt đầu", "Thủ thuật", "Nhân viên thực hiện", "Giờ bắt đầu", "Giờ kết thúc", "Loại giờ", "Thời gian thực hiện"],
            ...matchingAssignments
        ]);
    }
}



    console.log("Phân công, thống kê slot rảnh, ngoài giờ và kiểm tra thiếu thủ thuật hoàn tất.");
}

// Hàm hỗ trợ: Cộng thêm phút vào thời gian
function addMinutes(time, minutes) {
    if (typeof time !== "string") {
        throw new Error(`Tham số 'time' không hợp lệ: ${time}`);
    }
    if (typeof minutes !== "number" || isNaN(minutes)) {
        throw new Error(`Tham số 'minutes' không hợp lệ: ${minutes}`);
    }
    const date = new Date(`1970-01-01T${convertTo24HourFormat(time)}Z`);
    date.setMinutes(date.getMinutes() + minutes);
    return `${String(date.getUTCHours()).padStart(2, "0")}:${String(date.getUTCMinutes()).padStart(2, "0")}`;
}

// Hàm hỗ trợ: Chuyển thời gian thành số phút từ 00:00
function parseTime(time) {
    if (typeof time !== "string") {
        throw new Error(`Tham số 'time' không hợp lệ: ${time}`);
    }
    const [hours, minutes] = convertTo24HourFormat(time).split(":").map(Number);
    if (isNaN(hours) || isNaN(minutes)) {
        throw new Error(`Thời gian không hợp lệ: ${time}`);
    }
    return hours * 60 + minutes;
}

// Hàm hỗ trợ: Định dạng thời gian từ số phút
function formatTime(minutes) {
    console.log(`Đang định dạng thời gian từ số phút: ${minutes}`);
    if (typeof minutes === "number") {
        const hours = Math.floor(minutes / 60);
        const mins = minutes % 60;
        const formattedTime = `${String(hours).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
        console.log(`Thời gian đã định dạng: ${formattedTime}`);
        return formattedTime;
    }

    if (typeof minutes === "string") {
        console.log(`Thời gian đã là chuỗi: ${minutes}`);
        return minutes;
    }

    throw new Error(`Tham số 'minutes' không hợp lệ: ${minutes}`);
}

// Hàm hỗ trợ: Chuyển định dạng giờ từ 12 giờ (AM/PM) hoặc 24 giờ sang 24 giờ
function convertTo24HourFormat(time) {
    console.log(`Đang xử lý thời gian: "${time}"`);
    if (typeof time !== "string") {
        throw new Error(`Tham số 'time' không hợp lệ: ${time}`);
    }

    time = time.trim();
    if (time.includes("AM") || time.includes("PM")) {
        const [timePart, modifier] = time.split(" ");
        console.log(`Định dạng 12 giờ: timePart="${timePart}", modifier="${modifier}"`);
        let [hours, minutes, seconds] = timePart.split(":").map(Number);
        if (isNaN(hours) || isNaN(minutes)) {
            throw new Error(`Thời gian không hợp lệ: ${time}`);
        }
        if (seconds === undefined) {
            seconds = 0;
        }
        if (modifier === "PM" && hours !== 12) {
            hours += 12;
        } else if (modifier === "AM" && hours === 12) {
            hours = 0;
        }
        const formattedTime = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
        console.log(`Thời gian đã chuyển đổi: ${formattedTime}`);
        return formattedTime;
    } else if (time.includes(":")) {
        const parts = time.split(":").map(Number);
        if (parts.length === 2) {
            parts.push(0);
        }
        const [hours, minutes, seconds] = parts;
        if (isNaN(hours) || isNaN(minutes) || isNaN(seconds)) {
            throw new Error(`Thời gian không hợp lệ: ${time}`);
        }
        const formattedTime = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
        console.log(`Thời gian đã chuyển đổi: ${formattedTime}`);
        return formattedTime;
    } else {
        throw new Error(`Thời gian không hợp lệ: ${time}`);
    }
}

// Định nghĩa route kiểm tra
app.get("/", (req, res) => {
    res.send("Server Express đang chạy!");
});

// Định nghĩa route ghi dữ liệu lên Google Sheets
app.post("/write", async (req, res) => {
    try {
        const { sheetName, range, values } = req.body;
        await writeSheet(sheetName, range, values);
        res.send("Dữ liệu đã được ghi thành công!");
    } catch (error) {
        res.status(500).send(`Lỗi: ${error.message}`);
    }
});

// Route tạo các sheet mới từ cột A của sheet nhân viên
app.post("/createSheets", async (req, res) => {
    try {
        const { employeeSheetName, templateSheetName } = req.body;
        await createSheetsFromTemplate(employeeSheetName, templateSheetName);
        res.send("Các sheet đã được tạo thành công!");
    } catch (error) {
        res.status(500).send(`Lỗi: ${error.message}`);
    }
});

// Route xử lý chia giờ thực hiện thủ thuật
app.post("/allocateProcedures", async (req, res) => {
    try {
        await allocateProcedures();
        res.send("Chia giờ thực hiện thủ thuật thành công!");
    } catch (error) {
        res.status(500).send(`Lỗi: ${error.message}`);
    }
});

// Định nghĩa các route khác nếu cần

// Lắng nghe trên cổng 3000
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});

module.exports = { readSheet, writeSheet };
