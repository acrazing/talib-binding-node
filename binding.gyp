{
  "targets": [
    {
      "target_name": "talib_binding",
      "sources": [
        "src/talib-binding.generated.cc"
      ],
      "include_dirs": [
        "<!(node -e \"require('nan')\")"
      ],
      "link_settings": {
        "libraries": [
          "../ta-lib/c/lib/libta_libc_csr.a"
        ]
      }
    }
  ]
}