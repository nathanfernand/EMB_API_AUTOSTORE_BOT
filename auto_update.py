import main

AUTO_UPDATE_FILE = r"C:\base\BI_NATHAN\api_autostore_export.xlsx"


if __name__ == "__main__":
    main.os.makedirs(main.os.path.dirname(AUTO_UPDATE_FILE), exist_ok=True)
    if not main.os.path.exists(AUTO_UPDATE_FILE):
        main.pd.DataFrame().to_excel(AUTO_UPDATE_FILE, index=False)

    app = main.App()
    app.show_frame(main.ActionsPage)
    app.after(250, lambda: app.frames[main.ActionsPage].start_update_all(AUTO_UPDATE_FILE))
    app.mainloop()
