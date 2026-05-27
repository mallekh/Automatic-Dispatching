# Automatic-Dispatching
this work is only foir adult do not consult if you are under 18 we are not responsible for your bad behavior 

Quick test/run
-------------

You can generate training/export artifacts from the historical workbook using the provided script. The script now accepts tuning flags:

```bash
python train_dispatch_model.py --input data/Historique.xlsx --output-dir data_test \
	--max-passengers 4 --similarity-threshold 0.6
```

Adjust `--max-passengers` and `--similarity-threshold` to tune grouping behavior without editing code.
