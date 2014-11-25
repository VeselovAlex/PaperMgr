namespace PaperMgr.Entity
{
    public static class Labels
    {
        public enum LabelNames { None, Printed, Handwritten, Electronic }
        const string PRINTED = "печ.";
        const string HANDWRITTEN = "рук.";
        const string ELECTRONIC = "электр.";

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
