const openFileDialogBtn1 = document.getElementById("openFileDialogBtn1");
const openFileDialogBtn2 = document.getElementById("openFileDialogBtn2");
const filePath1Input = document.getElementById("filePath1Input");
const filePath2Input = document.getElementById("filePath2Input");
const selectForiegnKey1 = document.getElementById("selectForeignKey1");
const selectForiegnKey2 = document.getElementById("selectForeignKey2");
const fieldContainer1 = document.getElementById("fields1");
const fieldContainer2 = document.getElementById("fields2");
const fieldsCheckedContainer = document.getElementById(
  "fieldsCheckedContainer"
);
const exportBtn = document.getElementById("exportBtn");

let excel1Data = null;
let excel2Data = null;

let selectedFields = [];

function updateCheckedFieldsDisplay() {
  fieldsCheckedContainer.innerHTML = "";

  if (selectedFields.length === 0) {
    fieldsCheckedContainer.textContent = "No fields selected.";
  } else {
    selectedFields.forEach((field) => {
      const fieldDiv = document.createElement("div");
      fieldDiv.className = "px-1 border";
      fieldDiv.textContent = field;
      fieldsCheckedContainer.appendChild(fieldDiv);
    });
  }
}

function handleCheckboxChange(checkbox, header) {
  if (checkbox.checked) {
    selectedFields.push(header);
  } else {
    selectedFields = selectedFields.filter((field) => field !== header);
  }
  updateCheckedFieldsDisplay();
}

openFileDialogBtn1.addEventListener("click", async function () {
  const filePath = await window.electronAPI.openFileDialog("Excel Files", [
    "xlsx",
  ]);
  if (filePath) {
    filePath1Input.value = filePath;
    const jsonData = await window.electronAPI.readExcelFile(filePath);
    if (jsonData) {
      excel1Data = jsonData;

      const headers = jsonData[0];
      selectForiegnKey1.innerHTML = "";
      fieldContainer1.innerHTML = "";

      headers.forEach((header, index) => {
        const checkboxContainer = document.createElement("div");
        checkboxContainer.className = "form-check";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = index;
        checkbox.className = "form-check-input";
        checkbox.id = `checkbox1${index}`;
        checkbox.addEventListener("change", () =>
          handleCheckboxChange(checkbox, header)
        );

        const label = document.createElement("label");
        label.setAttribute("for", `checkbox1${index}`);
        label.textContent = header;
        label.className = "form-check-label";
        checkboxContainer.appendChild(checkbox);
        checkboxContainer.appendChild(label);

        fieldContainer1.appendChild(checkboxContainer);

        const option = document.createElement("option");
        option.value = index;
        option.textContent = header;
        selectForiegnKey1.appendChild(option);
      });
    }
  }
});

openFileDialogBtn2.addEventListener("click", async function () {
  const filePath = await window.electronAPI.openFileDialog("Excel Files", [
    "xlsx",
  ]);
  if (filePath) {
    filePath2Input.value = filePath;
    const jsonData = await window.electronAPI.readExcelFile(filePath);
    if (jsonData) {
      excel2Data = jsonData;

      const headers = jsonData[0];
      selectForiegnKey2.innerHTML = "";
      fieldContainer2.innerHTML = "";

      headers.forEach((header, index) => {
        const checkboxContainer = document.createElement("div");
        checkboxContainer.className = "form-check";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = index;
        checkbox.className = "form-check-input";
        checkbox.id = `checkbox2${index}`;
        checkbox.addEventListener("change", () =>
          handleCheckboxChange(checkbox, header)
        );

        const label = document.createElement("label");
        label.setAttribute("for", `checkbox2${index}`);
        label.textContent = header;
        label.className = "form-check-label";
        checkboxContainer.appendChild(checkbox);
        checkboxContainer.appendChild(label);

        fieldContainer2.appendChild(checkboxContainer);

        const option = document.createElement("option");
        option.value = index;
        option.textContent = header;
        selectForiegnKey2.appendChild(option);
      });
    }
  }
});

function displayFields(jsonData, fieldContainerId, selectForeignKey) {
  const fieldContainer = document.getElementById(fieldContainerId);
  const foreignKeySelect = document.getElementById(selectForeignKey);

  fieldContainer.innerHTML = "";
  foreignKeySelect.innerHTML = "";

  const headers = jsonData[0];

  headers.forEach((header, index) => {
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = index;
    fieldContainer.appendChild(checkbox);
    fieldContainer.appendChild(document.createTextNode(header));

    const option = document.createElement("option");
    option.value = index;
    option.textContent = header;
    foreignKeySelect.appendChild(option);
  });
}

exportBtn.addEventListener("click", async function () {
  // Kiểm tra xem dữ liệu Excel 1 và Excel 2 có tồn tại không
  if (!excel1Data || !excel2Data) {
    alert("Vui lòng chọn cả hai file Excel trước khi export.");
    return;
  }

  // Kiểm tra xem người dùng đã chọn trường nào chưa
  if (selectedFields.length === 0) {
    alert("Vui lòng chọn ít nhất một trường để export.");
    return;
  }

  // Lấy khóa ngoại từ các dropdown (ForeignKey1 và ForeignKey2)
  const foreignKey1 = selectForeignKey1.value;
  const foreignKey2 = selectForeignKey2.value;

  // Kiểm tra xem người dùng đã chọn khóa ngoại chưa
  if (!foreignKey1 || !foreignKey2) {
    alert("Vui lòng chọn khóa ngoại từ cả hai file Excel.");
    return;
  }

  // Chuẩn bị dữ liệu để gửi qua IPC
  const exportData = {
    excel1Data,
    excel2Data,
    selectedFields,
    foreignKey1,
    foreignKey2,
  };

  // Gửi dữ liệu qua IPC đến main process để xử lý việc export
  try {
    window.electronAPI.exportData(exportData);
  } catch (err) {
    console.error("Lỗi khi export:", err);
    alert("Đã xảy ra lỗi không mong muốn.");
  }
});
