# Automatic-Dispatching
this work is only foir adult do not consult if you are under 18 we are not responsible for your bad behavior 

## Retrain with new dataset
Use the command below to regenerate the model artifacts from `data/Historique.xlsx`:

```bash
python train_dispatch_model.py --input data/Historique.xlsx --output-dir data
```

This exports:
- `data/taxi_data_cleaned.csv`
- `data/final_course_dispatch_geographic.csv`
- `data/course_summary_geographic.csv`
- `data/model_training_report.json`
