def update_images(excel_path):
    src_path = 'D:\\iVolt_PC\\src_img_small'
    dest_path = 'D:\\iVolt_PC\\dest_img'

    copy_file = openpyxl.load_workbook(excel_path)
    data_list = copy_file.active

    copied_images = 0

    for row in range(2, data_list.max_row + 1):
        product_number = str(data_list.cell(row, 1).value)

        if not os.path.exists(dest_path):
            os.makedirs(dest_path)

        img_file = src_path + '\\' + product_number + '.jpg'

        if os.path.exists(img_file):
            print(f'Image {product_number} exist in {src_path} and was copied to {dest_path}')
            copied_images += 1
            shutil.copy2(img_file, os.path.join(dest_path, product_number + '.jpg'))
        else:
            print(f'{product_number}.jpg does not exist in the search directory.')

    print(f'Successfully copied {copied_images} images.')
