using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;

public static class Data
{

    public static List<Group> GroupsIndexes = new List<Group>();

    public static List<List<Shapes>> ShapesRanges = new List<List<Shapes>>();

    public static void AddToGroupList(string name, int index, int numberOfSheet)
    {
        bool isHave = false;
        for (int i =0; i< GroupsIndexes.Count; i++)
        {
            if (GroupsIndexes[i].Name == name && GroupsIndexes[i].Index == index)
            {
                isHave = true;
                break;
            }
        }
        if (!isHave)
        {
            GroupsIndexes.Add(new Group(name, index, numberOfSheet));
        }
    }

    public static void SaveGroupList()
    {
        JArray groups = new JArray();
        for (int i = 0; i < GroupsIndexes.Count; i++)
        {
            JObject group = new JObject();
            group.Add("name", GroupsIndexes[i].Name);
            group.Add("index", GroupsIndexes[i].Index);
            group.Add("numberOfSheet", GroupsIndexes[i].NumberOfSheet);
            groups.Add(group);
        }
        JObject jsonFile = new JObject();
        jsonFile.Add("data", groups);
        File.WriteAllText(@"D:\groupsIndexes.json", jsonFile.ToString());
    }

    public static void SaveShapesList()
    {
        JArray imagesPlaces = new JArray();
        for (int i = 0; i < ShapesRanges.Count; i++)
        {
            for (int j = 0; j < ShapesRanges[i].Count; j++)
            {
                JObject image = new JObject();
                image.Add("leftBorder", ShapesRanges[i][j].LeftBorder);
                image.Add("rightBorder", ShapesRanges[i][j].RightBorder);
                image.Add("topBorder", ShapesRanges[i][j].TopBorder);
                image.Add("bottomBorder", ShapesRanges[i][j].BottomBorder);
                image.Add("text", ShapesRanges[i][j].Text);
                imagesPlaces.Add(image);
            }
        }
        JObject jsonFile = new JObject();
        jsonFile.Add("data", imagesPlaces);
        File.WriteAllText(@"D:\imagesPlaces.json", jsonFile.ToString());
    }

    public class Group
    {
        private string name;
        private int index;
        private int numberOfSheet;

        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        public int Index
        {
            get
            {
                return index;
            }
            set
            {
                index = value;
            }
        }

        public int NumberOfSheet
        {
            get
            {
                return numberOfSheet;
            }
            set
            {
                numberOfSheet = value;
            }
        }

        public Group(string name = null, int index = 0, int numberOfSheet = 1)
        {
            Name = name;
            Index = index;
            NumberOfSheet = numberOfSheet;
        }
    }

    public class Shapes
    {
        private string text;
        private int leftBorder;
        private int rightBorder;
        private int topBorder;
        private int bottomBorder;

        public string Text
        {
            get
            {
                return text;
            }
            set
            {
                text = value;
            }
        }

        public int LeftBorder
        {
            get
            {
                return leftBorder;
            }
            set
            {
                leftBorder = value;
            }
        }

        public int RightBorder
        {
            get
            {
                return rightBorder;
            }
            set
            {
                rightBorder = value;
            }
        }

        public int TopBorder
        {
            get
            {
                return topBorder;
            }
            set
            {
                topBorder = value;
            }
        }

        public int BottomBorder
        {
            get
            {
                return bottomBorder;
            }
            set
            {
                bottomBorder = value;
            }
        }

        public Shapes(string text = null, int leftBorder = 0, int rightBorder = 0, int topBorder = 0, int bottomBorder = 0)
        {
            Text = text;
            LeftBorder = leftBorder;
            RightBorder = rightBorder;
            TopBorder = topBorder;
            BottomBorder = bottomBorder;
        }
    }
}
