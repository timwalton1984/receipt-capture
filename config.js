// Receipt Capture settings. Only these values need changing.
window.RECEIPT_CONFIG = {
  // Entra app registration "Receipt Capture" (Application (client) ID)
  clientId: "da1e9931-85b9-4ef9-8b7e-249f69efeb75",
  // CommitsDC tenant
  tenantId: "36f07eb5-3bb9-42c1-ab5c-9969359280ab",
  // Commits2 site document library (drive ID)
  driveId: "b!bG6Fe9O3AkuW2zQKThJRdjpFacXeRolOkwCcAzCdJkDoqrWpD4hFTYdYhTXS8kOi",
  // Folder path inside the library, above the "YYYY FY" folders
  basePath: ["Administration", "Finance", "Expense"],
  // Longest side of the saved photo, in pixels
  maxImageSize: 2400,
  jpegQuality: 0.85
};
