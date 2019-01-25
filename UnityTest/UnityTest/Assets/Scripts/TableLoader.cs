using System;
using System.Diagnostics;
using System.IO;
using UnityEngine;
using Debug = UnityEngine.Debug;

namespace FlatTable
{
	public enum LoadType
	{
		FilePath,
		ResourcePath
	}
	public class TableLoader
	{
		private static byte[] cacheBytes;

		public static byte[] GetByteArray(int length)
		{
			if (cacheBytes == null)
			{
				cacheBytes = new byte[1024];
			}

			if (length > cacheBytes.Length)
			{
				cacheBytes = new byte[length];
			}

			return cacheBytes;
		}

		public static string fileLoadPath;
		public static LoadType loadType = LoadType.FilePath;
		public static Func<string, byte[]> customLoader;

		public static void Load<T>() where T : TableBase, new()
		{
			T t = new T();

			if (customLoader != null)
			{
				byte[] bytes = customLoader(t.FileName);
				t.Decode(bytes);
				return;
			}
            
			if (loadType == LoadType.FilePath)
			{
				string fileName = t.FileName + ".bytes";
				string loadPath = Path.Combine(fileLoadPath, fileName);
				if (!File.Exists(loadPath))
				{
					Debug.LogError("无法根据路径载入配置文件： " + loadPath);
					t.Dispose();
					return;
				}

				using (FileStream fs = File.OpenRead(loadPath))
				{
					int byteLength = (int) fs.Length;
					fs.Position = 0;
					byte[] bytes = GetByteArray(byteLength);
					fs.Read(bytes, 0, byteLength);
					t.Decode(bytes);
				}
			}
			else if (loadType == LoadType.ResourcePath)
			{
                string loadPath = Path.Combine(fileLoadPath, t.FileName);
                TextAsset fileTextAsset = Resources.Load<TextAsset>(loadPath);
                if (fileTextAsset == null)
                {
	                Debug.LogError("无法在resource目录载入配置文件： " + loadPath);
	                return;
                }
                byte[] bytes = fileTextAsset.bytes;
                t.Decode(bytes);
			}
		}
	}
}