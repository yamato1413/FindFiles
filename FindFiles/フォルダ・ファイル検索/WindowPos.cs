namespace FindFiles
{
    class WindowPos
    {
        private int top, left, width, height;

        public const string DEFAULT_POSITION = "30 30 600 500";

        public WindowPos(int top, int left, int width, int height)
        {
            this.top = top < 0 ? 0 : top;
            this.left = left < 0 ? 0 : left;
            this.width = width;
            this.height = height;
        }

        public int Top { get { return top; } }
        public int Left { get { return left; } }
        public int Width { get { return width; } }
        public int Height { get { return height; } }
    }

}