"""Constants shared across GraphClient mixin modules."""

LARGE_FILE_THRESHOLD = 4 * 1024 * 1024      # 4 MB  — upload session threshold
UPLOAD_CHUNK_SIZE    = 5 * 1024 * 1024      # 5 MB  — upload chunk (must be multiple of 320 KB)

# Download / streaming constants
DOWNLOAD_TIMEOUT_SECONDS = 300              # 5 min — applies to ALL Graph API requests
STREAM_THRESHOLD_BYTES   = 10 * 1024 * 1024 # 10 MB — stream to disk above this size
STREAM_CHUNK_SIZE        = 5 * 1024 * 1024  # 5 MB  — download chunk size when streaming
