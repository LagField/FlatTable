using System.Collections;
using System.Collections.Generic;
using System.IO;
using FlatTable;
using UnityEngine;

public class LoadTest : MonoBehaviour
{
    void Start()
    {
//        LoadFromResourceFolder();
        LoadFromStreamingAssetsFolder();
    }

    private void LoadFromResourceFolder()
    {
        TableLoader.loadType = LoadType.ResourcePath;
        TableLoader.fileLoadPath = "Table/";
        TableLoader.Load<AnotherTestTable>();
        
        if (AnotherTestTable.ins == null)
        {
            return;
        }
        List<AnotherTestTable.Value> valueList = AnotherTestTable.ins.list;
        for (int i = 0; i < valueList.Count; i++)
        {
            Debug.Log(valueList[i].id);
            Debug.Log(valueList[i].hero_name);
            Debug.Log(valueList[i].speed);
            Debug.Log(valueList[i].damage);
            Debug.Log(valueList[i].is_lock);

            for (int j = 0; j < valueList[i].resource.Length; j++)
            {
                Debug.Log(valueList[i].resource[j]);
            }
        }
    }

    private void LoadFromStreamingAssetsFolder()
    {
        TableLoader.loadType = LoadType.FilePath;
        TableLoader.fileLoadPath = Application.streamingAssetsPath + "/Table";
        TableLoader.Load<AnotherTestTable>();
        
        if (AnotherTestTable.ins == null)
        {
            return;
        }
        List<AnotherTestTable.Value> valueList = AnotherTestTable.ins.list;
        for (int i = 0; i < valueList.Count; i++)
        {
            Debug.Log(valueList[i].id);
            Debug.Log(valueList[i].hero_name);
            Debug.Log(valueList[i].speed);
            Debug.Log(valueList[i].damage);
            Debug.Log(valueList[i].is_lock);

            for (int j = 0; j < valueList[i].resource.Length; j++)
            {
                Debug.Log(valueList[i].resource[j]);
            }
        }
    }
}