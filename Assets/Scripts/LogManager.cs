using UnityEngine;
using TMPro; // Import thư viện TextMeshPro
using System;
using UnityEngine.UI;

public class LogManager : MonoBehaviour
{
    [SerializeField] private TMPro.TextMeshProUGUI textPrefab; // Prefab của TextMeshPro

    // Hàm để thêm nội dung mới vào ScrollView
    public void AddLog(string message)
    {
        // check if the textPrefab's current line count is greater than 100
        if (textPrefab.textInfo.lineCount > 50)
        {
            // remove the first line
            textPrefab.text = textPrefab.text.Substring(textPrefab.text.IndexOf("\n") + 1);
        }
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // Thời gian hiện tại
        textPrefab.text += $"[{timestamp}] {message} \n";
    }
}
