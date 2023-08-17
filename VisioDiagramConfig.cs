namespace BlazorGraph
{
    public class VisioDiagramConfig
    {
        private double? _cardWidth;
        private double? _cardHeight;
        private double? _horizontalPageOffset;
        private double? _verticalPageOffset;
        private double? _initY;

        public double HeaderHeight { get; set; } = 0.5;
        public int CardsPerRow { get; set; } = 2;
        public int RowsPerPage { get; set; } = 2;
        public double PageWidth { get; set; } = 11;
        public double PageHeight { get; set; } = 8.5;
        public double HorizontalMargin { get; set; } = 0.2;
        public double VerticalMargin { get; set; } = 0.2;
        public double VerticalPageMargin { get; set; } = 0.5;
        public double HorizontalPageMargin { get; set; } = 0.5;
        public double MaxCardWidth { get; set; } = 2.0;
        public double MaxCardHeight { get; set; } = 1.5;

        public double AvailableDrawingWidth => PageWidth - (2 * HorizontalPageMargin);
        public double AvailableDrawingHeight => PageHeight - (2 * VerticalPageMargin);

        public double CardWidth
        {
            get => _cardWidth ??= Math.Min(MaxCardWidth, (PageWidth - HorizontalPageMargin * 2 - (CardsPerRow - 1) * HorizontalMargin) / CardsPerRow);
        }

        public double CardHeight
        {
            get => _cardHeight ??= Math.Min(MaxCardHeight, (PageHeight - VerticalPageMargin * 2 - (RowsPerPage - 1) * VerticalMargin) / RowsPerPage);
        }

        public double HorizontalPageOffset
        {
            get => _horizontalPageOffset ??= PageWidth - (CardsPerRow * CardWidth + (CardsPerRow - 1) * HorizontalMargin);
        }

        public double VerticalPageOffset
        {
            get => _verticalPageOffset ??= PageHeight - (RowsPerPage * CardHeight + (RowsPerPage - 1) * VerticalMargin);
        }

        public double InitY
        {
            get => _initY ??= PageHeight - VerticalPageMargin;
        }

        public void InvalidateCalculations()
        {
            _cardWidth = null;
            _cardHeight = null;
            _horizontalPageOffset = null;
            _verticalPageOffset = null;
            _initY = null;
        }
    }
}