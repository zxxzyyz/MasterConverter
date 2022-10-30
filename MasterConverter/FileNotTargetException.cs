using System;

public class FileNotTargetException : Exception
{
    public FileNotTargetException() {}

    public FileNotTargetException(string message) : base(message) {}

}
