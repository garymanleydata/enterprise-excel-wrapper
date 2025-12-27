select * from data_dictionary;

UPDATE data_dictionary
SET column_description = "['Month in which ', {'text': 'run ', 'bold': True, 'colour': 'red'}, ' took place.']"
WHERE column_name = 'total_distance';

"Month in which ", {'text': 'run ', 'bold': True, 'colour': 'red'}, " took place."