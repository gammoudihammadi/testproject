namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class ComboBoxOptions
    {
        public ComboBoxOptions()
        {
            IsUsedInFilter = true;
            ClickCheckAllAtStart = false;
            ClickUncheckAllAtStart = true;
            ClickCheckAllAfterSelection = false;
            ClickUncheckAllAfterSelection = false;
        }

        public ComboBoxOptions(string xpathId, string selectionValue) : this()
        {
            this.XpathId = xpathId;
            this.SelectionValue = selectionValue;
        }

        public ComboBoxOptions(string xpathId, string selectionValue, bool isUsedInFilter) : this(xpathId, selectionValue)
        {
            this.IsUsedInFilter = isUsedInFilter;
        }

        /// <summary>
        /// The ID used for XPath to find the multi select
        /// </summary>
        public string XpathId { get; set; }

        /// <summary>
        /// Le texte de la sélection multi select
        /// </summary>
        public string SelectionValue { get; set; }

        /// <summary>
        /// Définit si le multi select est utilisé en filtrage, et donc s'il faut attendre un chargement des résultats
        /// </summary>
        public bool IsUsedInFilter { get; set; }

        /// <summary>
        /// Indicateur de clique sur "check all" à l'ouverture du multi select
        /// </summary>
        public bool ClickCheckAllAtStart { get; set; }

        /// <summary>
        /// Indicateur de clique sur "uncheck all" à l'ouverture du multi select
        /// </summary>
        public bool ClickUncheckAllAtStart { get; set; }

        /// <summary>
        /// Indicateur de clique sur "check all" après avoir renseigner un filtre de sélection
        /// </summary>
        public bool ClickCheckAllAfterSelection { get; set; }

        /// Indicateur de clique sur "uncheck all" après avoir renseigner un filtre de sélection
        /// </summary>
        public bool ClickUncheckAllAfterSelection { get; set; }
    }
}
