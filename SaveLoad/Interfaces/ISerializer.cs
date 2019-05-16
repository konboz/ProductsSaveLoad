using System;
using System.Collections.Generic;
using System.Text;

namespace SaveLoad
{
    public interface ISerializer
    {
        bool SaveToFile(string fileName, Basket basket);
        Basket ReadFromFile(string fileName);
    }
}

