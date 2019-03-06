<?php

namespace main\controllers;

use Yii;
use app\models\Data;
use yii\web\Controller;
use yii\web\NotFoundHttpException;
use yii\filters\VerbFilter;
use yii\web\UploadedFile;
use PhpOffice\PhpSpreadsheet\IOFactory;

/* tabel data memiliki column
id
nama
nik
alamat
*/
class ImportSpreadsheetController extends Controller
{
    /**
     * {@inheritdoc}
     */
    public function behaviors()
    {
        return [
            'verbs' => [
                'class' => VerbFilter::className(),
                'actions' => [
                    'delete' => ['POST'],
                ],
            ],
        ];
    }

   
    public function actionImport()
    {
        $model = new Data();
        if ($model->load(Yii::$app->request->post())) {
            $file = UploadedFile::getInstance($model, 'files');

            try {
                $inputFileType = IOFactory::identify($file->tempName);

                $reader = IOFactory::createReader($inputFileType);
                $spreadsheet = $reader->load($file->tempName);
            } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
                die('Error loading file: '.$e->getMessage());
            }
            $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

       
            if ($this->saveImport($sheetData, $model)) {
                $this->redirect(['data/index']);
            } else {
                throw new NotFoundHttpException('salah database .');
            }
        }

        return $this->render('import', [
            'model' => $model,
        ]);
    }

  
    protected function findModel($id)
    {
        if (($model = TagihanKaryawan::findOne($id)) !== null) {
            return $model;
        }

        throw new NotFoundHttpException('The requested page does not exist.');
    }
     protected function saveImport($data, $model)
    {
        $rows = [];
        array_shift($data);
        foreach ($data as $key => $value) {
            if ($value['A'] != null) {
                $rows[] = array_merge_recursive([null], array_values($value)]);
            }
        }

        if (!empty($rows)) {
            try {
                return \Yii::$app->db->createCommand()->batchInsert(Data::tableName(), $model->attributes(), $rows)->execute();
            } catch (\yii\db\Exception $exception) {
                \Yii::warning('Kesalahan dalam eksekusi database.');
            }
        } else {
            return false;
        }
    }
    
    
}
