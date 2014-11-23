namespace PaperMgr.Entity
{
    public static class Labels
    {
        public enum LabelNames { None, Printed, Handwritten, Electronic }
        const string PRINTED = "Печатное издание";
        const string HANDWRITTEN = "Рукописное издание";
        const string ELECTRONIC = "Электронное издание";

        public static string ToString(Labels.LabelNames lbl)
        {
            switch (lbl)
            {
                case LabelNames.None:
                    return "";
                case LabelNames.Printed:
                    return PRINTED;
                case LabelNames.Handwritten:
                    return HANDWRITTEN;
                case LabelNames.Electronic:
                    return ELECTRONIC;
            }
            return "";
        }

    }
}
