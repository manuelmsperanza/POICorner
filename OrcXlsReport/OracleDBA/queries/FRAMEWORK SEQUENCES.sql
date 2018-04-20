--select sequence_name, min_value, max_value, increment_by, cache_size
select sequence_name, increment_by, cache_size
from user_sequences