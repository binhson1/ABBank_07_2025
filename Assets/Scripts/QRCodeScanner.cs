using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;
using ZXing;
using System.Threading.Tasks;
using TMPro;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class QRCodeScanner : MonoBehaviour
{
    public RawImage cameraPreview;
    public RenderTexture cameraTexture;
    public TextMeshProUGUI nameText, sexText, title1Text, title2Text, unitText, countText;
    public GameObject nameLabel, titleLabel;
    public GameObject standbyScreen;
    public TextMeshProUGUI tmpgui;
    public TMP_InputField manualInputField;
    public int scansPerSecond = 5;

    private WebCamTexture webcamTexture;
    private IBarcodeReader barcodeReader;

    private Dictionary<string, ExcelRow> guestDict = new Dictionary<string, ExcelRow>();
    private HashSet<string> checkedIn = new HashSet<string>();

    private string dataPath;
    private float standbyTimer;
    private float standbyTimeout = 60f;

    // NPOI objects
    private IWorkbook workbook;
    private ISheet worksheet;

    public LogManager logManager;

    void Start()
    {
        if (Display.displays.Length > 1)
        {
            Display.displays[1].Activate();
        }

        string filePath = Path.Combine(Path.GetDirectoryName(Application.dataPath), "Data");
        dataPath = Path.Combine(filePath, "Data.xlsx");

        LoadExcel();
        InitWebcam();
        barcodeReader = new BarcodeReader();
        standbyScreen.SetActive(false);

        StartCoroutine(ScanRoutine());
    }

    IEnumerator ScanRoutine()
    {
        while (true)
        {
            if (webcamTexture != null && webcamTexture.isPlaying)
            {
                tmpgui.text = "Camera OK - Đang quét...";
                bool scanned = false;

                try
                {
                    var data = webcamTexture.GetPixels32();
                    var width = webcamTexture.width;
                    var height = webcamTexture.height;

                    var result = barcodeReader.Decode(data, width, height);
                    if (result != null)
                    {
                        standbyTimer = 0f;
                        standbyScreen.SetActive(false);
                        ProcessQRCode(result.Text.Trim());
                        scanned = true;
                        tmpgui.text = $"Đã quét ID: {result.Text.Trim()}";
                        logManager?.AddLog($"Đã quét QR thành công: {result.Text.Trim()}");
                    }
                }
                catch (Exception e)
                {
                    Debug.LogWarning("Scan error: " + e.Message);
                    tmpgui.text = " Lỗi khi quét QR.";
                }

                if (scanned)
                {
                    yield return new WaitForSeconds(2f);
                }
            }
            else
            {
                tmpgui.text = " Không tìm thấy camera.";
            }

            yield return new WaitForSeconds(1f);
        }
    }

    void InitWebcam()
    {
        WebCamDevice[] devices = WebCamTexture.devices;
        if (devices.Length > 0)
        {
            webcamTexture = new WebCamTexture(devices[0].name, cameraTexture.width, cameraTexture.height);
            cameraPreview.texture = webcamTexture;
            webcamTexture.Play();
            cameraPreview.rectTransform.sizeDelta = new Vector2(webcamTexture.width, webcamTexture.height);
            cameraPreview.enabled = true;
            tmpgui.text = " Camera đã khởi động.";
        }
        else
        {
            tmpgui.text = " Không tìm thấy thiết bị camera.";
            Debug.LogError("No webcam found!");
        }
    }

    void LoadExcel()
    {
        using (FileStream file = new FileStream(dataPath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file);
            worksheet = workbook.GetSheet("DS Final");
        }

        guestDict.Clear();

        int row = 1; // NPOI index bắt đầu từ 0, dữ liệu bạn ở hàng 2 => index = 1
        while (true)
        {
            IRow excelRow = worksheet.GetRow(row);
            if (excelRow == null) break;

            string id = excelRow.GetCell(1)?.ToString(); // Cột B => index 1
            if (string.IsNullOrEmpty(id)) break;

            guestDict[id] = new ExcelRow
            {
                RowNumber = row,
                ID = id,
                Sex = excelRow.GetCell(3)?.ToString(),
                Name = excelRow.GetCell(4)?.ToString(),
                Unit = excelRow.GetCell(5)?.ToString(),
                Title1 = excelRow.GetCell(6)?.ToString(),
                Title2 = excelRow.GetCell(7)?.ToString()
            };
            row++;
        }

        tmpgui.text = $"Đã tải dữ liệu từ Excel ({guestDict.Count} người).";
    }

    public void ManualCheckIn()
    {
        string inputId = manualInputField.text.Trim();
        if (!string.IsNullOrEmpty(inputId))
        {
            logManager?.AddLog($"Check-in thủ công: {inputId}");
            tmpgui.text = $"Check-in thủ công: {inputId}";
            ProcessQRCode(inputId);
        }
        else
        {
            tmpgui.text = " Vui lòng nhập ID.";
        }
    }

    void ProcessQRCode(string id)
    {
        if (!guestDict.ContainsKey(id))
        {
            Debug.LogWarning("ID không tồn tại: " + id);
            tmpgui.text = $" ID không tồn tại: {id}";
            ShowInfo(null);
            logManager?.AddLog($" Không tìm thấy ID: {id}");
            return;
        }

        if (checkedIn.Contains(id))
        {
            tmpgui.text = $" Đã check-in trước đó: {id}";
            logManager?.AddLog($" Lặp lại ID đã check-in: {id}");
            checkedIn.Remove(id);
            // return;
        }

        ExcelRow guest = guestDict[id];
        ShowInfo(guest);

        IRow excelRow = worksheet.GetRow(guest.RowNumber);
        if (excelRow.GetCell(11) == null) excelRow.CreateCell(11);
        if (excelRow.GetCell(12) == null) excelRow.CreateCell(12);

        excelRow.GetCell(11).SetCellValue("Checked-in"); // Cột L => index 11
        excelRow.GetCell(12).SetCellValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")); // Cột M => index 12

        checkedIn.Add(id);

        SaveExcelAsync();
        countText.text = $"Số người đã checkin: {checkedIn.Count}";
        tmpgui.text = $"Check-in thành công: {guest.Name}";
        logManager?.AddLog($"Check-in thành công: {guest.Name} - ID: {id}");
    }

    async void SaveExcelAsync()
    {
        await Task.Run(() =>
        {
            using (FileStream file = new FileStream(dataPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }
        });
    }

    void ShowInfo(ExcelRow guest)
    {
        if (guest == null)
        {
            nameText.text = "";
            sexText.text = title1Text.text = title2Text.text = unitText.text = "";
        }
        else
        {
            nameText.text = guest.Name;
            sexText.text = guest.Sex;
            title1Text.text = guest.Title1;
            title2Text.text = guest.Title2;
            unitText.text = guest.Unit;
        }
        StartCoroutine(ForceRefreshLayoutNextFrame());
    }

    IEnumerator ForceRefreshLayoutNextFrame()
    {
        yield return null;
        LayoutRebuilder.ForceRebuildLayoutImmediate(nameLabel.GetComponent<RectTransform>());
        LayoutRebuilder.ForceRebuildLayoutImmediate(titleLabel.GetComponent<RectTransform>());
    }

    [Serializable]
    public class ExcelRow
    {
        public int RowNumber;
        public string ID, Sex, Name, Unit, Title1, Title2;
    }
}
