using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;
using ZXing;
using OfficeOpenXml;
using System.Threading.Tasks;
using TMPro;

public class QRCodeScanner : MonoBehaviour
{
    public RawImage cameraPreview;
    public RenderTexture cameraTexture;
    public TextMeshProUGUI nameText, sexText, title1Text, title2Text, unitText, countText;
    public GameObject nameLabel, titleLabel;
    public GameObject standbyScreen;
    public int scansPerSecond = 5;
    public string id = "";
    private WebCamTexture webcamTexture;
    private IBarcodeReader barcodeReader;
    private Dictionary<string, ExcelRow> guestDict = new Dictionary<string, ExcelRow>();
    private HashSet<string> checkedIn = new HashSet<string>();
    private string dataPath;
    private float scanTimer;
    private float standbyTimer;
    private float standbyTimeout = 60f;
    private ExcelPackage excelPackage;
    private ExcelWorksheet worksheet;

    void Start()
    {
        // active second display
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

        StartCoroutine(ScanRoutine()); // sử dụng Coroutine để quét QR
        // ProcessQRCode(id);
    }

    IEnumerator ScanRoutine()
    {
        while (true)
        {
            if (webcamTexture != null && webcamTexture.isPlaying)
            {
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
                    }
                }
                catch (Exception e)
                {
                    Debug.LogWarning("Scan error: " + e.Message);
                }
                if (scanned)
                {
                    yield return new WaitForSeconds(2f); // tránh quét trùng nhiều lần
                }
            }

            // standbyTimer += 1f / scansPerSecond;
            // if (standbyTimer >= standbyTimeout)
            // {
            //     standbyScreen.SetActive(true);
            // }

            yield return new WaitForSeconds(2f);
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
        }
        else
        {
            Debug.LogError("No webcam found!");
            return;
        }
    }

    void LoadExcel()
    {
        var fileInfo = new FileInfo(dataPath);
        excelPackage = new ExcelPackage(fileInfo);
        worksheet = excelPackage.Workbook.Worksheets["DS Final"];

        int row = 2;
        while (true)
        {
            var id = worksheet.Cells[row, 2].Text;
            if (string.IsNullOrEmpty(id)) break;

            guestDict[id] = new ExcelRow
            {
                RowNumber = row,
                ID = id,
                Sex = worksheet.Cells[row, 4].Text,
                Name = worksheet.Cells[row, 5].Text,
                Unit = worksheet.Cells[row, 6].Text,
                Title1 = worksheet.Cells[row, 7].Text,
                Title2 = worksheet.Cells[row, 8].Text
            };
            row++;
        }
    }

    void ProcessQRCode(string id)
    {
        if (!guestDict.ContainsKey(id))
        {
            Debug.LogWarning("ID không tồn tại: " + id);
            ShowInfo(null);
            return;
        }

        if (checkedIn.Contains(id))
        {
            Debug.Log("Đã checkin rồi");
            return;
        }

        ExcelRow guest = guestDict[id];
        ShowInfo(guest);
        // Mark checkin in Excel
        worksheet.Cells[guest.RowNumber, 12].Value = "Checked-in";
        checkedIn.Add(id);

        SaveExcelAsync();
        countText.text = $"Số người đã checkin: {checkedIn.Count}";
    }

    async void SaveExcelAsync()
    {
        await Task.Run(() =>
        {
            using (var stream = new MemoryStream())
            {
                excelPackage.SaveAs(stream);
                File.WriteAllBytes(dataPath, stream.ToArray());
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
        yield return null; // đợi 1 frame
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
