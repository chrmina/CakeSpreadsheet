<?php
namespace CakeSpreadsheet\View;

use Cake\Core\Exception\Exception;
use Cake\Event\EventManager;
use Cake\Network\Request;
use Cake\Network\Response;
use Cake\Utility\Inflector;
use Cake\View\View;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * @package  Cake.View
 */
class SpreadsheetView extends View
{
    /**
     * Excel layouts are located in the xlsx sub directory of `Layouts/`
     *
     * @var string
     */
    public $layoutPath = 'xlsx';

    /**
     * Excel views are always located in the 'xlsx' sub directory for a
     * controllers views.
     *
     * @var string
     */
    public $subDir = 'xlsx';

    /**
     * Spreadsheet instance
     *
     * @var Spreadsheet
     */
    public $Spreadsheet = null;

    /**
     * Constructor
     *
     * @param \Cake\Network\Request $request Request instance.
     * @param \Cake\Network\Response $response Response instance.
     * @param \Cake\Event\EventManager $eventManager EventManager instance.
     * @param array $viewOptions An array of view options
     */
    public function __construct(
        Request $request = null,
        Response $response = null,
        EventManager $eventManager = null,
        array $viewOptions = []
    ) {
        if (!empty($viewOptions['templatePath']) && $viewOptions['templatePath'] == '/xlsx') {
            $this->subDir = null;
        }

        parent::__construct($request, $response, $eventManager, $viewOptions);

        if (isset($viewOptions['name']) && $viewOptions['name'] == 'Error') {
            $this->subDir = null;
            $this->layoutPath = null;
            $response->type('html');

            return;
        }

        if ($response && $response instanceof Response) {
            $response->type('xlsx');
        }

        $this->Spreadsheet = new Spreadsheet();
    }

    /**
     * Magic accessor for helpers. Backward compatibility for PHPExcel property
     *
     * @param string $name Name of the attribute to get.
     *
     * @return mixed
     */
    public function __get($name)
    {
        if ($name === 'PhpExcel') {
            deprecationWarning('The `PhpExcel` property is deprecated. Use SpreadsheetView::$Spreadsheet instead.');

            return $this->Spreadsheet;
        }

        return parent::__get($name);
    }

    /**
     * Render method
     *
     * @param string $view The view being rendered.
     * @param string $layout The layout being rendered.
     * @return string The rendered view.
     */
    public function render($view = null, $layout = null)
    {
        $content = parent::render($view, $layout);
        if ($this->response->type() == 'text/html') {
            return $content;
        }

        $this->Blocks->set('content', $this->output());
        $this->response->download($this->getFilename());

        return $this->Blocks->get('content');
    }

    /**
     * Generates the binary excel data
     *
     * @return string
     */
    protected function output()
    {
        ob_start();

        $writer = new Xlsx($this->Spreadsheet);

        $writer->setPreCalculateFormulas(false);
        $writer->setIncludeCharts(true);
        $writer->save('php://output');

        $output = ob_get_clean();

        return $output;
    }

    /**
     * Gets the filename
     *
     * @return string filename
     */
    public function getFilename()
    {
        if (isset($this->viewVars['_filename'])) {
            return $this->viewVars['_filename'] . '.xlsx';
        }

        return Inflector::slug(str_replace('.xlsx', '', $this->request->url)) . '.xlsx';
    }
}
