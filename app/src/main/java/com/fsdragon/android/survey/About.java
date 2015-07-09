package com.fsdragon.android.survey;

import android.app.AlertDialog;
import android.app.Dialog;
import android.app.DialogFragment;
import android.content.DialogInterface;
import android.os.Bundle;

/**
 * Created by Semion on 09/07/2015.
 */
public class About extends DialogFragment {
    @Override
    public Dialog onCreateDialog(Bundle savedInstanceState){
        AlertDialog.Builder builder = new AlertDialog.Builder(getActivity());
        builder.setTitle(R.string.About);
        builder.setMessage("Created By: Semion Faifman\nJuly 2015");
        builder.setPositiveButton(R.string.Close, new DialogInterface.OnClickListener() {
            @Override
            public void onClick(DialogInterface dialog, int which) {
                dismiss();
            }
        });
        Dialog dialog = builder.create();
        return dialog;
    }
}
