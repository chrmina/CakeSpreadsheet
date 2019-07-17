<?php
namespace CakeSpreadsheet\Test\TestCase\View;

use CakeSpreadsheet\View\SpreadsheetView;
use Cake\Network\Request;
use Cake\Network\Response;
use Cake\TestSuite\TestCase;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * CsvViewTest
 */
class SpreadsheetViewTest extends TestCase
{

    public $fixtures = ['core.Articles', 'core.Authors'];

    /**
     * setup callback
     *
     * @return void
     */
    public function setUp()
    {
        $this->request = new Request([
            'params' => [
                'plugin' => null,
                'controller' => 'posts',
                'action' => 'index',
                '_ext' => null,
                'pass' => []
            ],
            'url' => '/posts',
        ]);
        $this->response = new Response();

        $this->View = new SpreadsheetView($this->request, $this->response);
    }

    public function tearDown()
    {
        unset($this->View);
    }

    /**
     * testRender
     *
     */
    public function testConstruct()
    {
        $result = $this->View->response->type();
        $this->assertEquals('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', $result);
        $this->assertTrue($this->View->Spreadsheet instanceof Spreadsheet);
    }

    public function testRender()
    {
        $this->View->name = $this->View->viewPath = 'Posts';

        $output = $this->View->render('index');
        $this->assertSame('504b030414', bin2hex(substr($output, 0, 5)));

        $result = $this->View->Spreadsheet->getActiveSheet()->getCellByColumnAndRow(0, 0)->getValue();
        $this->assertEquals('Test string', $result);
    }

    public function testGetFilename()
    {
        $result = $this->View->getFilename();
        $this->assertEquals('posts.xlsx', $result);

        $this->View->set('_filename', 'filename test');
        $result = $this->View->getFilename();
        $this->assertEquals('filename test.xlsx', $result);
    }
}
