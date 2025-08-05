using UnityEngine;
using TMPro;
using System;
using UnityEngine.UI;

public class LogManager : MonoBehaviour
{
    [SerializeField] private GameObject textPrefab;
    [SerializeField] private Transform content;
    [SerializeField] private ScrollRect scrollRect;


    public void AddLog(string message)
    {

        GameObject newTextObject = Instantiate(textPrefab, content);


        TextMeshProUGUI textComponent = newTextObject.GetComponent<TextMeshProUGUI>();
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        textComponent.text = $"[{timestamp}] {message}";


        LayoutRebuilder.ForceRebuildLayoutImmediate(content.GetComponent<RectTransform>());


        Canvas.ForceUpdateCanvases();
        scrollRect.verticalNormalizedPosition = 0f;
    }
}
