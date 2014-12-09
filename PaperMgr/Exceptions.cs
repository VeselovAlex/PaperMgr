using System;
namespace PaperMgr.Exceptions
{
    [Serializable]
    public class YearOutOfRangeException : Exception
    {
        public YearOutOfRangeException() { }
        public YearOutOfRangeException(string message) : base(message) { }
        public YearOutOfRangeException(string message, Exception inner) : base(message, inner) { }
        protected YearOutOfRangeException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }

    [Serializable]
    public class PageCountNumberException : Exception
    {
        public PageCountNumberException() { }
        public PageCountNumberException(string message) : base(message) { }
        public PageCountNumberException(string message, Exception inner) : base(message, inner) { }
        protected PageCountNumberException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }

    [Serializable]
    public class AuthorMismatchException : Exception
    {
        public AuthorMismatchException() { }
        public AuthorMismatchException(string message) : base(message) { }
        public AuthorMismatchException(string message, Exception inner) : base(message, inner) { }
        protected AuthorMismatchException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }
}
